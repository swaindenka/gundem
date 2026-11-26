"""Okunan belgeden gündem tablosu için gerekli alanları çıkarır."""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from .model import Cell, Document, Table

_TR_MAP = str.maketrans({
    "ı": "i", "İ": "I", "ş": "s", "Ş": "S", "ğ": "g", "Ğ": "G",
    "ü": "u", "Ü": "U", "ö": "o", "Ö": "O", "ç": "c", "Ç": "C",
    "â": "a", "Â": "A", "î": "i", "Î": "I", "û": "u", "Û": "U",
    "̇": "",
})


def tr_upper(s: str) -> str:
    """Türkçe kurallarına uygun büyük harf (i -> İ, ı -> I)."""
    return s.replace("i", "İ").replace("ı", "I").upper()


def norm(s: str) -> str:
    """Anahtar kelime karşılaştırması için: Türkçe harfleri sadeleştirip büyük harfe çevirir."""
    return s.translate(_TR_MAP).upper()


@dataclass
class CaseData:
    kaynak: str = ""
    sira_no: str = ""
    esas: str = ""
    mahkeme: str = ""
    daire: str = ""
    daire_karar: str = ""
    konu: str = ""
    alt_bolum: str = ""
    uyarilar: List[str] = field(default_factory=list)


# --------------------------------------------------------------------------
# Üst tablo
# --------------------------------------------------------------------------
def classify_label(text: str) -> Optional[str]:
    n = norm(text)
    n = re.sub(r"\s+", " ", n).strip()
    if not n or len(n) > 40:
        return None
    if "SIRA" in n:
        return "sira"
    if "ESAS" in n and "MAHKEME" in n:
        return "esas_mahkeme"
    if "ESAS" in n:
        return "esas"
    if "MAHKEME" in n:
        return "mahkeme"
    if "DAIRE" in n:
        return "daire"
    if "KONU" in n:
        return "konu"
    if "KARAR" in n:
        return "karar"
    return None


_DATE_RE = re.compile(r"\d{1,2}\s*[./]\s*\d{1,2}\s*[./]\s*\d{4}")


def format_karar(text: str) -> str:
    """'13.12.2027 tarihli ve 2027/9000 E. ...' metnini şablondaki gibi iki satıra böler."""
    t = re.sub(r"[ \t]+", " ", text).strip()
    if "\n" in t:
        return "\n".join(l.strip() for l in t.split("\n") if l.strip())
    m = re.search(r"(tarihli\s+ve)\s+", t, flags=re.IGNORECASE)
    if m:
        return t[:m.end(1)] + "\n" + t[m.end():]
    return t


def split_daire(text: str) -> Tuple[str, str]:
    """'3. Hukuk Dairesinin 13.12.2027 tarihli ve ... kararı' -> (daire, karar)."""
    t = text.strip()
    lines = [l.strip() for l in t.split("\n") if l.strip()]
    if len(lines) >= 2 and not _DATE_RE.search(lines[0]) and _DATE_RE.search(lines[1]):
        return lines[0], format_karar("\n".join(lines[1:]))
    m = _DATE_RE.search(t)
    if not m:
        return t, ""
    if m.start() == 0:
        return "", format_karar(t)
    name = t[:m.start()].strip().rstrip(",;:-")
    name = re.sub(r"[’'`]?\s*n[ıi]n$", "", name).strip()
    return name, format_karar(t[m.start():])


def _grid_rows(table: Table) -> List[List[Tuple[int, int, Cell]]]:
    out = []
    for row in table.rows:
        col = 0
        cells = []
        for c in row:
            cells.append((col, col + max(c.span, 1), c))
            col += max(c.span, 1)
        out.append(cells)
    return out


def _clean_value(text: str) -> str:
    lines = [re.sub(r"[ \t]+", " ", l).strip() for l in text.split("\n")]
    return "\n".join(l for l in lines if l)


def _from_header_layout(table: Table) -> Optional[Dict[str, str]]:
    """Şablondaki gibi: başlık satırı (ESAS NO-MAHKEME | DAİRE KARARI | KONUSU) + altında değerler."""
    grid = _grid_rows(table)
    for hi, row in enumerate(grid):
        labels = [(s, e, classify_label(c.text)) for s, e, c in row]
        kinds = {k for _, _, k in labels if k}
        if len(kinds) < 2:
            continue
        cols = [(s, e, k) for s, e, k in labels if k]
        values: Dict[str, List[str]] = {k: [] for _, _, k in cols}
        for row2 in grid[hi + 1:]:
            for s, e, c in row2:
                if c.vcontinue:
                    continue
                v = _clean_value(c.text)
                if not v:
                    continue
                # hücrenin en çok örtüştüğü başlık sütununu bul
                best, best_ov = None, 0
                for hs, he, k in cols:
                    ov = min(e, he) - max(s, hs)
                    if ov > best_ov:
                        best, best_ov = k, ov
                if best:
                    values[best].append(v)
        return _combine(values)
    return None


def _from_label_layout(table: Table) -> Optional[Dict[str, str]]:
    """Etiket | : | değer biçimindeki satırlar."""
    values: Dict[str, List[str]] = {}
    for row in table.rows:
        texts = [_clean_value(c.text) for c in row if not c.vcontinue]
        texts = [t for t in texts if t and t.strip(" :") != ""]
        if not texts:
            continue
        first = texts[0]
        rest = texts[1:]
        if ":" in first and not rest:
            label, _, val = first.partition(":")
            kind = classify_label(label)
            rest = [val.strip()] if val.strip() else []
        else:
            label = first
            if ":" in first and classify_label(first.partition(":")[0]):
                label, _, extra = first.partition(":")
                if extra.strip():
                    rest = [extra.strip()] + rest
            kind = classify_label(label)
        if not kind or not rest:
            continue
        val = "\n".join(r.lstrip(":").strip() for r in rest).strip()
        if val:
            values.setdefault(kind, []).append(val)
    if len(values) < 2:
        return None
    return _combine(values)


def _combine(values: Dict[str, List[str]]) -> Dict[str, str]:
    out: Dict[str, str] = {}
    if values.get("esas_mahkeme"):
        v = values["esas_mahkeme"]
        out["esas"] = v[0]
        out["mahkeme"] = "\n".join(v[1:])
    if values.get("esas"):
        out["esas"] = "\n".join(values["esas"])
    if values.get("mahkeme"):
        out["mahkeme"] = "\n".join(values["mahkeme"])
    if values.get("daire"):
        v = values["daire"]
        if len(v) >= 2:
            out["daire"] = v[0]
            out["daire_karar"] = format_karar("\n".join(v[1:]))
        else:
            out["daire"], out["daire_karar"] = split_daire(v[0])
    if values.get("karar") and not out.get("daire_karar"):
        out["daire_karar"] = format_karar("\n".join(values["karar"]))
    if values.get("konu"):
        seen = []
        for v in values["konu"]:
            if v not in seen:
                seen.append(v)
        out["konu"] = "\n".join(seen)
    if values.get("sira"):
        out["sira_no"] = values["sira"][0]
    return out


def _from_paragraphs(lines: List[str]) -> Optional[Dict[str, str]]:
    values: Dict[str, List[str]] = {}
    for line in lines[:60]:
        if ":" not in line:
            continue
        label, _, val = line.partition(":")
        kind = classify_label(label)
        if kind and val.strip() and kind not in values:
            values[kind] = [_clean_value(val)]
    if len(values) < 2:
        return None
    return _combine(values)


def extract_header(doc: Document) -> Dict[str, str]:
    for table in doc.tables[:5]:
        for fn in (_from_header_layout, _from_label_layout):
            res = fn(table)
            if res and len(res) >= 2:
                return res
    return _from_paragraphs(doc.lines()) or {}


# --------------------------------------------------------------------------
# Alt bölüm (******* işaretinden sonrası)
# --------------------------------------------------------------------------
def _is_star_line(line: str) -> bool:
    s = line.replace(" ", "").replace("\t", "")
    return len(s) >= 3 and set(s) <= set("*")


_STRONG_START = re.compile(r"KARARI(NI)?\s+VEREN")
_HAKIM_WORD = re.compile(r"HAKIMI\b|HAKIMLERI\b")


def _clean_lines(lines: List[str]) -> List[str]:
    out = [l.strip() for l in lines]
    return [re.sub(r"[ ]{2,}", "  ", l) for l in out if l]


def _is_judge_line(n: str) -> bool:
    return bool(_MEMBER_RE.match(n) or _STRONG_START.search(n) or _HAKIM_WORD.search(n))


def _judge_blocks(lines: List[str]) -> List[List[str]]:
    """Metindeki ardışık hakim satırı bloklarını bulur.

    Blok, 'kararı veren' ya da 'BAŞKAN/ÜYE' satırıyla başlar; hakim satırları sürdükçe devam eder.
    '... hakimi' ile biten satırdan sonra gelen tek satırlık isim de bloğa dahil edilir.
    """
    blocks: List[List[str]] = []
    i, n = 0, len(lines)
    while i < n:
        nl = norm(lines[i])
        if not (_STRONG_START.search(nl) or _MEMBER_RE.match(nl)) or _is_star_line(lines[i]):
            i += 1
            continue
        j = i
        block: List[str] = []
        while j < n and not _is_star_line(lines[j]):
            nj = norm(lines[j])
            if _is_judge_line(nj):
                block.append(lines[j])
            elif block and re.search(r"HAKIMI\s*:?$", norm(block[-1])) and len(lines[j]) <= 80:
                block.append(lines[j])  # ayrı satırdaki ilk derece hakimi adı
            else:
                break
            j += 1
        has_judge = any(_MEMBER_RE.match(norm(l)) or _HAKIM_WORD.search(norm(l)) for l in block)
        if has_judge:
            blocks.append(block)
        i = max(j, i + 1)
    return blocks


def find_alt_bolum(doc: Document) -> Tuple[str, str]:
    """Hakim bilgilerini bulur. (metin, kaynak) döndürür; kaynak: yildiz / emsal / genel / ''.

    1. '*******' işaretinin altı (hakim satırı içeriyorsa)
    2. Yoksa 'Emsal' bölümünden sonra gelen hakim satırları
    3. Yoksa belgedeki son hakim satırı bloğu
    """
    lines = _clean_lines(doc.lines())
    stars = [i for i, l in enumerate(lines) if _is_star_line(l)]
    fallback: Optional[List[str]] = None
    for idx in reversed(stars):
        end = next((j for j in stars if j > idx), len(lines))
        section = lines[idx + 1:end]
        if not section:
            continue
        if fallback is None:
            fallback = section
        if _judge_blocks(section):
            return "\n".join(section), "yildiz"

    emsal = [i for i, l in enumerate(lines) if "EMSAL" in norm(l)]
    for idx in reversed(emsal):
        blocks = _judge_blocks(lines[idx + 1:])
        if blocks:
            return "\n".join(blocks[0]), "emsal"

    blocks = _judge_blocks(lines)
    if blocks:
        return "\n".join(blocks[-1]), "genel"
    if fallback:
        return "\n".join(fallback), "yildiz"
    return "", ""


def extract_alt_bolum(doc: Document) -> str:
    return find_alt_bolum(doc)[0]


# Alt bölüm satır türleri
BASLIK = "baslik"   # "İlk BAM kararını veren ... hakimleri:" (tam genişlik satır)
UYE = "uye"         # "BAŞKAN : Ad Soyad (sicil)" (değerlendirme kutulu satır)
TEK = "tek"         # "İlk kararı veren ilk derece mahkemesi ... hakimi Ad Soyad" (değerlendirme kutulu satır)

_MEMBER_RE = re.compile(
    r"^\s*((?:KIDEMLI\s+)?(?:BASKAN(?:\s*V(?:EKILI)?\.?)?|UYE|HAKIM|KATIP))\b\s*[:\-]?\s*(.*)$"
)


def parse_alt_bolum(text: str) -> List[Tuple[str, str]]:
    raw = [l.strip() for l in text.split("\n") if l.strip()]
    items: List[Tuple[str, str]] = []
    i = 0
    while i < len(raw):
        line = raw[i]
        n = norm(line)
        m = _MEMBER_RE.match(n)
        if m:
            label_len = len(m.group(1))
            label = tr_upper(line[:label_len].strip())
            rest = line[label_len:].strip().lstrip(":-").strip()
            rest = re.sub(r"\s+", " ", rest)
            items.append((UYE, "%s\t: %s" % (label, rest)))
        elif line.rstrip().endswith(":") or "HAKIMLERI" in n:
            items.append((BASLIK, re.sub(r"\s+", " ", line)))
        else:
            text_line = re.sub(r"\s+", " ", line)
            # "... hakimi" ile biten satırdan sonra isim ayrı satırdaysa birleştir
            if re.search(r"HAKIMI\s*:?$", n) and i + 1 < len(raw):
                nxt = norm(raw[i + 1])
                if not _MEMBER_RE.match(nxt) and not raw[i + 1].rstrip().endswith(":") and "HAKIMLERI" not in nxt:
                    text_line = text_line.rstrip(":").rstrip() + " " + re.sub(r"\s+", " ", raw[i + 1])
                    i += 1
            items.append((TEK, text_line))
        i += 1
    return items


# --------------------------------------------------------------------------
def extract_case(doc: Document, kaynak: str = "") -> CaseData:
    data = CaseData(kaynak=kaynak)
    h = extract_header(doc)
    data.sira_no = h.get("sira_no", "")
    data.esas = h.get("esas", "")
    data.mahkeme = h.get("mahkeme", "")
    data.daire = h.get("daire", "")
    data.daire_karar = h.get("daire_karar", "")
    data.konu = h.get("konu", "")
    data.alt_bolum, alt_kaynak = find_alt_bolum(doc)
    if not h:
        data.uyarilar.append("Belgenin üstündeki tablo bulunamadı; esas/mahkeme/daire/konu alanlarını elle doldurun.")
    else:
        eksik = [ad for ad, k in (("Esas", "esas"), ("Mahkeme", "mahkeme"), ("Daire", "daire"), ("Konu", "konu"))
                 if not h.get(k)]
        if eksik:
            data.uyarilar.append("Bulunamayan alanlar: %s" % ", ".join(eksik))
    if not data.alt_bolum:
        data.uyarilar.append("Hakim bilgileri bulunamadı ('*******' altında ve 'Emsal' bölümünden sonra bakıldı).")
    elif alt_kaynak == "emsal":
        data.uyarilar.append("Hakim bilgileri '*******' altında yoktu; 'Emsal' bölümünden sonra bulundu. Kontrol edin.")
    elif alt_kaynak == "genel":
        data.uyarilar.append("Hakim bilgileri '*******' altında yoktu; belgenin içinden bulundu. Kontrol edin.")
    return data


def looks_complete(doc: Document) -> bool:
    """Yerleşik .doc okuyucunun sonucunun kullanılabilir olup olmadığını kabaca denetler."""
    return bool(extract_header(doc)) and bool(extract_alt_bolum(doc))
