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
    alt_kaynak: str = ""  # hakim bilgilerinin bulunduğu yer: yildiz / emsal / genel / ''
    alt_konum: float = 1.0  # hakim listesinin belgedeki konumu (0 = baş, 1 = son); uzun belgede
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
# Alt bölüm (******* işaretinden sonrası): hakim bilgileri
# --------------------------------------------------------------------------
# Alt bölüm satır türleri
BASLIK = "baslik"   # "İlk BAM kararını veren ... hakimleri:" (tam genişlik satır)
UYE = "uye"         # "BAŞKAN : Ad Soyad (sicil)" (değerlendirme kutulu satır)
TEK = "tek"         # "İlk kararı veren ilk derece mahkemesi ... hakimi Ad Soyad" (değerlendirme kutulu satır)


def _is_star_line(line: str) -> bool:
    """'*******' işareti. Metin içindeki '***' / '* * *' gibi kısa ayraçlar sayılmaz."""
    s = line.replace(" ", "").replace("\t", "")
    return len(s) >= 5 and set(s) <= set("*")


_STRONG_START = re.compile(r"KARARI(NI)?\s+VEREN|HUKMU\s+VEREN")
_HAKIMI = re.compile(r"HAKIMI(?![A-Z])")
_HAKIMLERI = re.compile(r"HAKIMLERI(?![A-Z])")
_TETKIK = re.compile(r"TETKIK\s+HAKIM")

# Görev adları (norm() uygulanmış metin üzerinde). Uzun olanlar önce denenir.
_ROL = (r"(?:KIDEMLI\s+UYE|BASKAN\s*VEKILI|BASKAN\s*V\.?(?![A-Z])|DAIRE\s+BASKANI|UYE\s+HAKIM"
        r"|BASKAN|UYE|HAKIM|REIS)")
_ROL_BAS = re.compile(r"^(%s)(?![A-Z])\s*[:\-–]?\s*" % _ROL)
_ROL_ICINDE = re.compile(r"(?<![A-Z])(%s)(?![A-Z])" % _ROL)
_MADDE = re.compile(r"^\s*(?:[-–•*·]+|\d{1,2}\s*[.)\-–]|[A-Za-z]\))\s*")
_SICIL = re.compile(r"(?:SICIL(?:\s*NO)?\s*[:.]?\s*)?\(?\s*(\d{3,})\s*\)?")
_SICIL_TEK = re.compile(r"^(?:SICIL(?:\s*NO)?\s*[:.]?\s*)?\(?\s*\d{3,}\s*\)?$")
_DEGERLENDIRME = {"COK", "IYI", "ORTA", "ZAYIF", "NOTSUZ", "ACIKLAMA", "PUAN", "NOT"}


def _rol_etiketi(n_rol: str) -> str:
    r = re.sub(r"\s+", " ", n_rol).strip().rstrip(".")
    if r.startswith("KIDEMLI"):
        return "KIDEMLİ ÜYE"
    if r.replace(" ", "") == "BASKANVEKILI":
        return "BAŞKANVEKİLİ"
    if r in ("BASKAN V", "BASKANV"):
        return "BAŞKAN V."
    if r == "DAIRE BASKANI":
        return "DAİRE BAŞKANI"
    return {"BASKAN": "BAŞKAN", "UYE": "ÜYE", "UYE HAKIM": "ÜYE", "HAKIM": "HAKİM", "REIS": "REİS"}.get(r, r)


def degerlendirme_mi(text: str) -> bool:
    """Şablondaki 'Çok iyi / İyi / Orta / Zayıf / Notsuz / AÇIKLAMA:' kutusu metni mi?"""
    kelimeler = re.findall(r"[A-Z]+", norm(text))
    return bool(kelimeler) and all(k in _DEGERLENDIRME for k in kelimeler)


def _hazirla(line: str) -> str:
    line = line.replace("̇", "").replace("ß", "ss").replace("\t", " ")
    line = _MADDE.sub("", line, count=1) if _MADDE.match(line) and not _SICIL_TEK.match(norm(line.strip())) else line
    return re.sub(r"\s+", " ", line).strip()


def _isim_sicil(rest: str) -> Tuple[str, str, str]:
    """'Ad Soyad (12345)' / 'Ad Soyad 12345' / 'Ad Soyad Sicil No: 12345' -> (isim, sicil, ek)."""
    rest = rest.strip()
    n = norm(rest)
    son = None
    for m in _SICIL.finditer(n):
        son = m
    if son is None:
        return _temizle_isim(rest), "", ""
    isim = _temizle_isim(rest[:son.start()])
    ek = rest[son.end():].strip(" -–,;:/")
    return isim, son.group(1), ek


def _temizle_isim(t: str) -> str:
    t = re.sub(r"\(\s*\)", " ", t)
    t = re.sub(r"\s+", " ", t)
    return t.strip(" -–,;:/(")


def _isim_gibi(line: str) -> bool:
    """'Ad SOYAD', 'Mehmet Ali Yılmaz 12345' gibi kısa, büyük harfle başlayan sözcüklerden oluşan satır."""
    t = line.strip()
    if not t or len(t) > 70:
        return False
    t = re.sub(r"\(?\s*\d{3,}\s*\)?", " ", t)
    t = re.sub(r"(?i)sicil\s*(no)?\s*[:.]?", " ", t)
    words = re.findall(r"[^\s\-–,;:/()]+", t)
    if not 2 <= len(words) <= 6:
        return False
    return all(w[0].isupper() for w in words) and all(any(c.isalpha() for c in w) for w in words)


def _uye_metni(rol: str, isim: str, sicil: str, ek: str = "") -> str:
    t = "%s\t: %s" % (rol, isim)
    if sicil:
        t += "  (%s)" % sicil
    if ek:
        t += " " + ek
    return t


class _Kalem:
    def __init__(self, tur: str, metin: str = "", rol: str = "", isim: str = "", sicil: str = "", ek: str = ""):
        self.tur, self.metin, self.rol, self.isim, self.sicil, self.ek = tur, metin, rol, isim, sicil, ek

    def bitti_mi(self) -> bool:
        return bool(self.isim)

    def metne(self) -> str:
        if self.tur == UYE:
            return _uye_metni(self.rol, self.isim, self.sicil, self.ek)
        return self.metin


def _uye_ayristir(line: str) -> Optional[_Kalem]:
    """Görev içeren satırı (BAŞKAN : Ad Soyad 123 / Ad Soyad (Üye) 123) ayrıştırır."""
    n = norm(line)
    if _STRONG_START.search(n) or _HAKIMLERI.search(n) or _TETKIK.search(n):
        return None
    m = _ROL_BAS.match(n)
    if m:
        rol = _rol_etiketi(m.group(1))
        isim, sicil, ek = _isim_sicil(line[m.end():])
        return _Kalem(UYE, rol=rol, isim=isim, sicil=sicil, ek=ek)
    if len(line) <= 90:
        m = _ROL_ICINDE.search(n)
        if m:
            kalan = line[:m.start()] + " " + line[m.end():]
            if _isim_gibi(_temizle_isim(kalan)) or _SICIL.search(norm(kalan)):
                isim, sicil, ek = _isim_sicil(kalan)
                if isim:
                    return _Kalem(UYE, rol=_rol_etiketi(m.group(1)), isim=isim, sicil=sicil, ek=ek)
    return None


def _kalemler(lines: List[str]) -> List[_Kalem]:
    kalemler: List[_Kalem] = []
    for raw in lines:
        line = _hazirla(raw)
        if not line or line.strip(" :-–.") == "" or _is_star_line(line) or degerlendirme_mi(line):
            continue
        n = norm(line)
        son = kalemler[-1] if kalemler else None

        uye = _uye_ayristir(line)
        if uye is not None:
            kalemler.append(uye)
            continue
        if _SICIL_TEK.match(n):
            # Ayrı satırdaki sicil numarası: önceki hakime ekle
            sicil = re.search(r"\d{3,}", n).group(0)
            if son is not None and son.tur == UYE and not son.sicil:
                son.sicil = sicil
            elif son is not None and son.tur == TEK:
                son.metin += " " + sicil
            continue
        if son is not None and son.tur == UYE and not son.isim and _isim_gibi(line):
            son.isim, son.sicil, son.ek = _isim_sicil(line)
            continue
        if (son is not None and son.tur == TEK and re.search(r"HAKIMI\s*:?$", norm(son.metin))
                and _isim_gibi(line)):
            son.metin = son.metin.rstrip(" :") + " " + line
            continue
        if line.endswith(":") or _HAKIMLERI.search(n):
            kalemler.append(_Kalem(BASLIK, metin=line))
        else:
            kalemler.append(_Kalem(TEK, metin=line))
    # Hakim adı içermeyen ve ardından üye satırları gelen satır bir başlıktır
    for i, k in enumerate(kalemler[:-1]):
        if k.tur == TEK and not _HAKIMI.search(norm(k.metin)) and kalemler[i + 1].tur == UYE:
            k.tur = BASLIK
    return kalemler


def parse_alt_bolum(text: str) -> List[Tuple[str, str]]:
    return [(k.tur, k.metne()) for k in _kalemler(text.split("\n"))]


def _hakim_iceriyor(kalemler: List[_Kalem], siki: bool = False) -> bool:
    """Kalemlerde gerçekten hakim bilgisi var mı?

    siki=True (belge gövdesinde arama yaparken): isimli üye satırı ya da '... hakimi Ad Soyad'
    biçiminde isim içeren satır gerekir; 'mahkeme hakimi gerekçesinde ...' gibi cümleler sayılmaz.
    """
    for k in kalemler:
        if k.tur == UYE and k.isim:
            return True
        n = norm(k.metin)
        m = _HAKIMI.search(n)
        if k.tur == TEK and m:
            if not siki:
                return True
            if _isim_gibi(k.metin[m.end():]):
                return True
        if not siki and k.tur == BASLIK and (_HAKIMLERI.search(n) or _STRONG_START.search(n)):
            return True
    return False


def _clean_lines(lines: List[str]) -> List[str]:
    out = [l.strip() for l in lines]
    return [re.sub(r"[ ]{2,}", "  ", l) for l in out if l]


def _hakim_satiri_mi(line: str) -> bool:
    n = norm(line)
    if _TETKIK.search(n):
        return False  # raporu yazan tetkik hakiminin imzası hakim listesine dahil değildir
    return bool(_STRONG_START.search(n) or _HAKIMLERI.search(n) or _HAKIMI.search(n)
                or _uye_ayristir(_hazirla(line)) is not None or _ROL_BAS.match(norm(_hazirla(line))))


def _bloklar(lines: List[str], siki: bool = True) -> List[Tuple[int, int]]:
    """Ardışık hakim satırı bloklarını (başlangıç, bitiş) indeksleriyle bulur."""
    blocks: List[Tuple[int, int]] = []
    i, n = 0, len(lines)
    while i < n:
        if _is_star_line(lines[i]) or not _hakim_satiri_mi(lines[i]):
            i += 1
            continue
        j = i
        while j < n and not _is_star_line(lines[j]) and not _TETKIK.search(norm(lines[j])):
            h = _hazirla(lines[j])
            if (_hakim_satiri_mi(lines[j]) or _isim_gibi(h) or _SICIL_TEK.match(norm(h))
                    or degerlendirme_mi(h) or not h.strip(" :-–.")):
                j += 1
            else:
                break
        if _hakim_iceriyor(_kalemler(lines[i:j]), siki=siki):
            blocks.append((i, j))
        i = max(j, i + 1)
    return blocks


_KUME_ARALIGI = 4  # bloklar arasında en fazla bu kadar alakasız satır varsa aynı listedir


def _kumeler(lines: List[str], siki: bool = True) -> List[Tuple[int, int]]:
    """Birbirine yakın hakim bloklarını tek liste (küme) olarak birleştirir.

    Ör. ilk derece hakimi satırı, bir boş/ara satır, ardından BAM hakimleri: tek küme.
    """
    kumeler: List[Tuple[int, int]] = []
    for a, b in _bloklar(lines, siki):
        if kumeler:
            pa, pb = kumeler[-1]
            ara = lines[pb:a]
            if (a - pb <= _KUME_ARALIGI and not any(_is_star_line(l) for l in ara)
                    and not any(_TETKIK.search(norm(l)) for l in ara)):
                kumeler[-1] = (pa, b)
                continue
        kumeler.append((a, b))
    return kumeler


def _judge_blocks(lines: List[str]) -> List[List[str]]:
    return [lines[a:b] for a, b in _bloklar(lines)]


def _tetkikte_kes(section: List[str]) -> List[str]:
    """Hakim listesinin altındaki 'Tetkik Hakimi ...' imzasını ve sonrasını atar."""
    for i, l in enumerate(section):
        if _TETKIK.search(norm(l)):
            return section[:i]
    return section


_UZUN_BOLUM = 40  # yıldız altı bundan uzunsa, ortadaki bir ayraçtan başlayıp metnin geri kalanını içeriyordur


def find_alt_bolum(doc: Document) -> Tuple[str, str]:
    metin, kaynak, _ = _alt_bolum_bul(doc)
    return metin, kaynak


def _alt_bolum_bul(doc: Document) -> Tuple[str, str, float]:
    """Hakim bilgilerini bulur. (metin, kaynak) döndürür; kaynak: yildiz / emsal / genel / ''.

    Hakim isimleri ve sicilleri belgenin SON sayfalarında yer alır. Bu yüzden adaylar arasından
    belgenin sonuna en yakın olan seçilir:
      - '*******' altındaki bölümler (hakim bilgisi içeriyorsa)
      - belge içindeki hakim listeleri (isim içeren BAŞKAN/ÜYE satırları vb.)
    Ortadaki bir yıldız ayracı ya da metin içinde alıntılanan eski bir karar, sondaki gerçek
    listenin önüne geçemez.
    """
    lines = _clean_lines(doc.lines(satir_bazli=True, atla=degerlendirme_mi))
    stars = [i for i, l in enumerate(lines) if _is_star_line(l)]

    # (başlangıç, bitiş, kaynak) adayları; indeksler `lines` içindeki mutlak konumlardır
    adaylar: List[Tuple[int, int, str]] = []
    fallback: Optional[Tuple[int, int]] = None
    for k, idx in enumerate(stars):
        end = stars[k + 1] if k + 1 < len(stars) else len(lines)
        section = _tetkikte_kes(lines[idx + 1:end])
        if not section:
            continue
        a, b = idx + 1, idx + 1 + len(section)
        fallback = (a, b)
        if not _hakim_iceriyor(_kalemler(section)):
            continue
        if len(section) > _UZUN_BOLUM:
            # Uzun bölüm: yalnızca sondaki hakim listesini al
            kumeler = _kumeler(section, siki=False)
            if not kumeler:
                continue
            ka, kb = kumeler[-1]
            a, b = idx + 1 + ka, idx + 1 + kb
        adaylar.append((a, b, "yildiz"))

    emsal = [i for i, l in enumerate(lines) if "EMSAL" in norm(l)]
    tetkik = [i for i, l in enumerate(lines) if _TETKIK.search(norm(l))]
    for a, b in _kumeler(lines, siki=True):
        if any(sa <= a < sb for sa, sb, _ in adaylar):
            continue  # zaten bir yıldız bölümünün içinde
        if any(0 < a - t <= 3 for t in tetkik):
            continue  # 'Tetkik Hakimi' imzasının altındaki isim: hakim listesi değil
        kaynak = "emsal" if any(e < a for e in emsal) else "genel"
        adaylar.append((a, b, kaynak))

    def konum(b: int) -> float:
        # kısa belgelerde konum anlamsız; uyarı yalnızca uzun belgeler için
        return b / len(lines) if len(lines) > 150 else 1.0

    if adaylar:
        a, b, kaynak = max(adaylar, key=lambda x: (x[0], x[2] == "yildiz"))
        return "\n".join(lines[a:b]), kaynak, konum(b)
    if fallback:
        return "\n".join(lines[fallback[0]:fallback[1]]), "yildiz", konum(fallback[1])
    return "", "", 1.0


def alt_bolum_duzenle(text: str) -> str:
    """Ham hakim satırlarını tabloya yazılacak biçime getirir (tekrar uygulanabilir)."""
    return "\n".join(t for _, t in parse_alt_bolum(text))


def extract_alt_bolum(doc: Document) -> str:
    return find_alt_bolum(doc)[0]


# --------------------------------------------------------------------------
def uyarilari_hesapla(data: CaseData) -> List[str]:
    uyarilar: List[str] = []
    alanlar = (("Esas", "esas"), ("Mahkeme", "mahkeme"), ("Daire", "daire"), ("Konu", "konu"))
    eksik = [ad for ad, k in alanlar if not getattr(data, k)]
    if len(eksik) == len(alanlar):
        uyarilar.append("Belgenin üstündeki tablo bulunamadı; esas/mahkeme/daire/konu alanlarını elle doldurun.")
    elif eksik:
        uyarilar.append("Bulunamayan alanlar: %s" % ", ".join(eksik))
    if not data.alt_bolum:
        uyarilar.append("Hakim bilgileri bulunamadı ('*******' altında ve 'Emsal' bölümünden sonra bakıldı).")
    elif data.alt_kaynak == "emsal":
        uyarilar.append("Hakim bilgileri '*******' altında yoktu; 'Emsal' bölümünden sonra bulundu. Kontrol edin.")
    elif data.alt_kaynak == "genel":
        uyarilar.append("Hakim bilgileri '*******' altında yoktu; belgenin içinden bulundu. Kontrol edin.")
    if data.alt_bolum and data.alt_konum < 0.8:
        uyarilar.append("DİKKAT: Hakim bilgileri belgenin sonundan değil, ortalarından (yaklaşık %%%d) alındı. "
                        "Son sayfalardaki listeyle karşılaştırın." % round(data.alt_konum * 100))
    return uyarilar


def extract_case(doc: Document, kaynak: str = "") -> CaseData:
    data = CaseData(kaynak=kaynak)
    h = extract_header(doc)
    data.sira_no = h.get("sira_no", "")
    data.esas = h.get("esas", "")
    data.mahkeme = h.get("mahkeme", "")
    data.daire = h.get("daire", "")
    data.daire_karar = h.get("daire_karar", "")
    data.konu = h.get("konu", "")
    ham, data.alt_kaynak, data.alt_konum = _alt_bolum_bul(doc)
    # Kutuda tabloya yazılacak düzenlenmiş hali gösterilir (BAŞKAN : Ad Soyad  (sicil));
    # böylece görev/isim/sicil doğru ayrıldı mı kullanıcı kaydetmeden görebilir.
    data.alt_bolum = alt_bolum_duzenle(ham)
    data.uyarilar = uyarilari_hesapla(data)
    return data


def looks_complete(doc: Document, bolum: bool = False) -> bool:
    """Yerleşik .doc okuyucunun sonucunun kullanılabilir olup olmadığını kabaca denetler.

    Bölümlere ayrılmış dosyalarda (bolum=True) üst tablo ilk bölümde, hakimler son bölümde
    olduğundan ikisinden birinin bulunması yeterlidir.
    """
    tablo = bool(extract_header(doc))
    hakim = bool(extract_alt_bolum(doc))
    if bolum:
        # ara bölümlerde tablo da hakim de olmayabilir: metin düzgün okunduysa yeterli
        return tablo or hakim or len("".join(doc.lines()).strip()) > 200
    return tablo and hakim
