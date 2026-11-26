"""Kaynak Word belgelerini (.docx ve .doc) ortak modele okur."""
from __future__ import annotations

import glob
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
from typing import Callable, List, Optional

from lxml import etree

from .model import Cell, Document, Table, sembol_karakteri

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}


def _q(tag: str) -> str:
    return "{%s}%s" % (W, tag)


class ReadError(Exception):
    pass


# --------------------------------------------------------------------------
# .docx
# --------------------------------------------------------------------------
MC_FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"
_ATLA = {_q("delText"), _q("instrText"), _q("pPr"), _q("rPr"), MC_FALLBACK}


def _pua_cevir(t: str) -> str:
    """Symbol yazı tipindeki karakterler U+F000-U+F0FF aralığında gelir; yıldızı '*' yap."""
    if not any("\uf000" <= c <= "\uf0ff" for c in t):
        return t
    return "".join(sembol_karakteri(ord(c)) if "\uf000" <= c <= "\uf0ff" else c for c in t)


def _para_text(p) -> str:
    parts: List[str] = []

    def walk(el):
        for child in el:
            tag = child.tag
            if not isinstance(tag, str) or tag in _ATLA:
                # mc:Fallback: metin kutularının eski (VML) kopyası; okunursa metin iki kez gelir
                continue
            if tag == _q("t"):
                parts.append(_pua_cevir(child.text or ""))
            elif tag == _q("sym"):
                # Ekle > Simge ile eklenen karakter (ör. Symbol yazı tipinde yıldız: F02A)
                try:
                    parts.append(sembol_karakteri(int(child.get(_q("char")) or "0", 16)))
                except ValueError:
                    parts.append("\ufffc")
            elif tag == _q("tab"):
                parts.append("\t")
            elif tag in (_q("br"), _q("cr")):
                parts.append("\n")
            elif tag == _q("noBreakHyphen"):
                parts.append("-")
            elif tag == _q("txbxContent"):
                # metin kutusu: içindeki paragrafları ayrı satır olarak ekle
                for tp in child.iter(_q("p")):
                    parts.append("\n" + _para_text(tp))
                parts.append("\n")
                continue
            elif tag == _q("p"):
                continue
            walk(child)

    walk(p)
    return "".join(parts).replace("\xa0", " ")


def _iter_block_children(parent):
    """Gövde/hücre altındaki p ve tbl öğelerini, sdt/customXml sarmalarını açarak verir."""
    for child in parent:
        tag = child.tag
        if tag in (_q("p"), _q("tbl")):
            yield child
        elif tag == _q("sdt"):
            content = child.find(_q("sdtContent"))
            if content is not None:
                yield from _iter_block_children(content)
        elif tag == _q("customXml"):
            yield from _iter_block_children(child)


def _cell_text(tc) -> str:
    lines: List[str] = []
    for el in _iter_block_children(tc):
        if el.tag == _q("p"):
            lines.append(_para_text(el))
        else:  # iç içe tablo: metni satır satır ekle
            lines.extend(l for l in _table_from_xml(el).lines() if l.strip())
    return "\n".join(lines)


def _direct(parent, tag: str):
    """`parent` altındaki doğrudan `tag` öğelerini (sdt/customXml sarmaları açılarak) verir."""
    for child in parent:
        if child.tag == _q(tag):
            yield child
        elif child.tag == _q("sdt"):
            content = child.find(_q("sdtContent"))
            if content is not None:
                yield from _direct(content, tag)
        elif child.tag == _q("customXml"):
            yield from _direct(child, tag)


def _table_from_xml(tbl) -> Table:
    table = Table()
    for tr in _direct(tbl, "tr"):
        row: List[Cell] = []
        for tc in _direct(tr, "tc"):
            span = 1
            vcont = False
            tcpr = tc.find(_q("tcPr"))
            if tcpr is not None:
                gs = tcpr.find(_q("gridSpan"))
                if gs is not None:
                    try:
                        span = int(gs.get(_q("val")))
                    except (TypeError, ValueError):
                        span = 1
                vm = tcpr.find(_q("vMerge"))
                if vm is not None and vm.get(_q("val")) in (None, "continue"):
                    vcont = True
            row.append(Cell(text=_cell_text(tc), span=span, vcontinue=vcont))
        # gridBefore: satır başında atlanan ızgara sütunları
        trpr = tr.find(_q("trPr"))
        if trpr is not None:
            gb = trpr.find(_q("gridBefore"))
            if gb is not None:
                try:
                    n = int(gb.get(_q("val")))
                    if n > 0:
                        row.insert(0, Cell(text="", span=n))
                except (TypeError, ValueError):
                    pass
        table.rows.append(row)
    return table


def read_docx(path: str) -> Document:
    try:
        with zipfile.ZipFile(path) as z:
            xml = z.read("word/document.xml")
    except (zipfile.BadZipFile, KeyError) as e:
        raise ReadError("Geçerli bir .docx dosyası değil: %s" % e)
    root = etree.fromstring(xml)
    body = root.find(_q("body"))
    doc = Document()
    if body is None:
        return doc
    for el in _iter_block_children(body):
        if el.tag == _q("p"):
            doc.blocks.append(_para_text(el))
        else:
            doc.blocks.append(_table_from_xml(el))
    return doc


# --------------------------------------------------------------------------
# .doc / .rtf -> .docx dönüştürücüler (yedek yol)
# --------------------------------------------------------------------------
def _convert_with_word(src: str, out_dir: str) -> Optional[str]:
    if sys.platform != "win32":
        return None
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except ImportError:
        return None
    pythoncom.CoInitialize()
    word = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(
            os.path.abspath(src), ConfirmConversions=False, ReadOnly=True,
            AddToRecentFiles=False, Visible=False,
        )
        out = os.path.join(out_dir, "donusturulen.docx")
        doc.SaveAs2(out, FileFormat=16)  # wdFormatDocumentDefault (.docx)
        doc.Close(False)
        return out if os.path.exists(out) else None
    except Exception:
        return None
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def _find_soffice() -> Optional[str]:
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    if sys.platform == "win32":
        pats = []
        for env in ("ProgramFiles", "ProgramFiles(x86)", "ProgramW6432"):
            base = os.environ.get(env)
            if base:
                pats.append(os.path.join(base, "LibreOffice*", "program", "soffice.exe"))
                pats.append(os.path.join(base, "OpenOffice*", "program", "soffice.exe"))
        for pat in pats:
            hits = glob.glob(pat)
            if hits:
                return hits[0]
    return None


def _convert_with_soffice(src: str, out_dir: str) -> Optional[str]:
    exe = _find_soffice()
    if not exe:
        return None
    try:
        kwargs = {}
        if sys.platform == "win32":
            kwargs["creationflags"] = 0x08000000  # CREATE_NO_WINDOW
        subprocess.run(
            [exe, "--headless", "--convert-to", "docx", "--outdir", out_dir, src],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=120, **kwargs,
        )
    except Exception:
        return None
    base = os.path.splitext(os.path.basename(src))[0] + ".docx"
    out = os.path.join(out_dir, base)
    return out if os.path.exists(out) else None


def _convert_and_read(path: str, log: Callable[[str], None]) -> Optional[Document]:
    with tempfile.TemporaryDirectory() as tmp:
        # Dosya adındaki Türkçe karakterler bazı dönüştürücüleri bozabildiği için kopyala
        ext = os.path.splitext(path)[1].lower() or ".doc"
        src = os.path.join(tmp, "kaynak" + ext)
        shutil.copyfile(path, src)
        for name, fn in (("Microsoft Word", _convert_with_word), ("LibreOffice", _convert_with_soffice)):
            out = fn(src, tmp)
            if out:
                log("%s ile .docx'e dönüştürüldü." % name)
                return read_docx(out)
    return None


def _sniff(path: str) -> str:
    with open(path, "rb") as f:
        head = f.read(8)
    if head.startswith(b"PK"):
        return "docx"
    if head.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
        return "doc"
    if head.startswith(b"{\\rtf"):
        return "rtf"
    return "unknown"


def read_document(path: str, log: Callable[[str], None] = lambda s: None,
                  is_good: Callable[[Document], bool] = lambda d: True) -> Document:
    """Belgeyi biçimine göre okur.

    .doc dosyaları önce yerleşik okuyucu ile okunur; sonuç `is_good` kontrolünden
    geçmezse Microsoft Word ya da LibreOffice ile dönüştürülerek tekrar denenir.
    """
    kind = _sniff(path)
    if kind == "docx":
        return read_docx(path)
    if kind == "doc":
        from .doc_binary import DocReadError, read_doc
        doc: Optional[Document] = None
        try:
            doc = read_doc(path)
            if is_good(doc):
                return doc
            log("Yerleşik .doc okuyucu eksik sonuç verdi, dönüştürücü deneniyor...")
        except DocReadError as e:
            log("Yerleşik .doc okuyucu: %s" % e)
        except Exception as e:  # beklenmeyen biçim
            log("Yerleşik .doc okuyucu hata verdi (%s), dönüştürücü deneniyor..." % e)
        converted = _convert_and_read(path, log)
        if converted is not None:
            return converted
        if doc is not None:
            return doc
        raise ReadError("Belge okunamadı. Dosyayı Word'de açıp .docx olarak kaydetmeyi deneyin.")
    if kind == "rtf":
        converted = _convert_and_read(path, log)
        if converted is not None:
            return converted
        raise ReadError(
            "Dosya aslında RTF biçiminde ve dönüştürmek için Word/LibreOffice bulunamadı. "
            "Dosyayı Word'de açıp .docx olarak kaydedin."
        )
    raise ReadError("Desteklenmeyen dosya biçimi (yalnızca .docx ve .doc).")
