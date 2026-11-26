"""Bölümlere ayrılmış (birinci bölüm, ikinci bölüm, ...) dosyaları tek dava olarak birleştirir.

Uzun dosyalar "2025-123 birinci bölüm.docx", "2025-123 2. bölüm.doc" gibi adlarla parçalara
ayrılır. Dosya adındaki bölüm ifadesi çıkarılınca geriye kalan ad aynıysa parçalar aynı davaya
aittir. Üst tablo bilgileri ilk bölümden, hakim bilgileri son bölümden alınır.
"""
from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from .extract import CaseData, norm, uyarilari_hesapla

SIRA_SOZCUKLERI = {
    "birinci": 1, "ikinci": 2, "ucuncu": 3, "dorduncu": 4, "besinci": 5,
    "altinci": 6, "yedinci": 7, "sekizinci": 8, "dokuzuncu": 9, "onuncu": 10,
}
ROMA = {"i": 1, "ii": 2, "iii": 3, "iv": 4, "v": 5, "vi": 6, "vii": 7, "viii": 8, "ix": 9, "x": 10}
_BOLUM = r"(?:bolum|bolumu|kisim|kisimi|parca|parcasi|part)"
_AYRAC = r"[\s_\-\.]*"
_EK = r"(?:'?\s*(?:inci|nci|ıncı|uncu|ncu|unci|nci))?"

_KALIPLAR = [
    # "birinci bölüm", "ikinci_kisim"
    re.compile(r"(?<![a-z])(%s)%s%s(?![a-z])" % ("|".join(SIRA_SOZCUKLERI), _AYRAC, _BOLUM)),
    # "1. bölüm", "1.bolum", "1 bölüm", "1'inci bölüm"
    re.compile(r"(?<![\d/])(\d{1,2})\.?%s%s%s(?![a-z])" % (_EK, _AYRAC, _BOLUM)),
    # "bölüm 2", "part2", "bölüm-3"
    re.compile(r"(?<![a-z])%s%s(\d{1,2})(?!\d)" % (_BOLUM, _AYRAC)),
    # "II. bölüm"
    re.compile(r"(?<![a-z])(%s)\.?%s%s(?![a-z])" % ("|".join(sorted(ROMA, key=len, reverse=True)), _AYRAC, _BOLUM)),
]


def bolum_bilgisi(path: str) -> Tuple[str, Optional[int]]:
    """Dosya adından (grup anahtarı, bölüm numarası) çıkarır. Bölüm yoksa numara None olur."""
    ad = os.path.splitext(os.path.basename(path))[0]
    n = norm(ad).lower()
    for kalip in _KALIPLAR:
        m = kalip.search(n)
        if not m:
            continue
        deger = m.group(1)
        if deger.isdigit():
            no = int(deger)
        elif deger in SIRA_SOZCUKLERI:
            no = SIRA_SOZCUKLERI[deger]
        else:
            no = ROMA[deger]
        kalan = n[:m.start()] + " " + n[m.end():]
        anahtar = re.sub(r"[\s_\-\.,;()\[\]]+", " ", kalan).strip()
        return anahtar, no
    return "", None


@dataclass
class Parca:
    no: int
    yol: str
    veri: CaseData
    kayit: List[str] = field(default_factory=list)  # okuma sırasında oluşan bilgi mesajları


HEADER_ALANLARI = ("sira_no", "esas", "mahkeme", "daire", "daire_karar", "konu")
TUM_ALANLAR = HEADER_ALANLARI + ("alt_bolum",)
_KAYNAK_ONCELIK = {"yildiz": 0, "emsal": 1, "genel": 2}


def birlestir(parcalar: List[Parca]) -> CaseData:
    """Bölümleri tek davaya birleştirir.

    - Üst tablo alanları: numarası en küçük olan ve o alanı içeren bölümden
    - Hakim bilgileri: hakim bilgisi bulunan en son bölümden
    """
    sirali = sorted(parcalar, key=lambda p: p.no)
    sonuc = CaseData(kaynak=sirali[0].yol)
    for alan in HEADER_ALANLARI:
        for p in sirali:
            deger = getattr(p.veri, alan)
            if deger:
                setattr(sonuc, alan, deger)
                break
    adaylar = [p for p in sirali if p.veri.alt_bolum]
    if adaylar:
        # Hakim isimleri belgenin sonunda olur: en son bölüm öncelikli, eşitlikte '*******' altı
        en_iyi = min(adaylar, key=lambda p: (-p.no, _KAYNAK_ONCELIK.get(p.veri.alt_kaynak, 3)))
        sonuc.alt_bolum = en_iyi.veri.alt_bolum
        sonuc.alt_kaynak = en_iyi.veri.alt_kaynak
        sonuc.alt_konum = en_iyi.veri.alt_konum

    uyarilar: List[str] = []
    for p in sirali:
        uyarilar.extend("%d. bölüm: %s" % (p.no, k) for k in p.kayit)
    numaralar = [p.no for p in sirali]
    if len(sirali) == 1:
        uyarilar.append("Dosya adına göre bu %d. bölüm; aynı adlı diğer bölümleri eklediğinizde "
                        "otomatik olarak bu davayla birleştirilir." % numaralar[0])
    else:
        uyarilar.append("%d bölüm birleştirildi (%s)." % (len(sirali), ", ".join("%d." % n for n in numaralar)))
    eksik = [n for n in range(1, max(numaralar) + 1) if n not in numaralar]
    if eksik:
        uyarilar.append("Eksik bölüm olabilir: %s bölüm eklenmedi." % ", ".join("%d." % n for n in eksik))
    tekrar = sorted({n for n in numaralar if numaralar.count(n) > 1})
    if tekrar:
        uyarilar.append("Aynı bölüm birden fazla eklendi: %s" % ", ".join("%d." % n for n in tekrar))
    uyarilar.extend(uyarilari_hesapla(sonuc))
    sonuc.uyarilar = uyarilar
    return sonuc


@dataclass
class Dava:
    """Gündemdeki bir dava: tek dosya ya da birleştirilmiş bölümler."""
    veri: CaseData
    grup: str = ""
    parcalar: List[Parca] = field(default_factory=list)
    otomatik: Dict[str, str] = field(default_factory=dict)  # son otomatik doldurulan değerler

    @classmethod
    def tek_dosya(cls, veri: CaseData) -> "Dava":
        return cls(veri=veri, otomatik={a: getattr(veri, a) for a in TUM_ALANLAR})

    @classmethod
    def bolumden(cls, grup: str, parca: Parca) -> "Dava":
        d = cls(veri=CaseData(), grup=grup)
        d.parca_ekle(parca)
        return d

    def parca_ekle(self, parca: Parca) -> None:
        """Yeni bölümü ekleyip alanları yeniden hesaplar; kullanıcının elle değiştirdiği alanlar korunur."""
        self.parcalar.append(parca)
        yeni = birlestir(self.parcalar)
        for alan in TUM_ALANLAR:
            mevcut = getattr(self.veri, alan)
            if mevcut == self.otomatik.get(alan, ""):
                setattr(self.veri, alan, getattr(yeni, alan))
            self.otomatik[alan] = getattr(yeni, alan)
        self.veri.kaynak = yeni.kaynak
        self.veri.alt_kaynak = yeni.alt_kaynak
        self.veri.uyarilar = yeni.uyarilar

    @property
    def dosyalar(self) -> List[str]:
        if self.parcalar:
            return [p.yol for p in sorted(self.parcalar, key=lambda p: p.no)]
        return [self.veri.kaynak]

    def etiket(self) -> str:
        ad = os.path.basename(self.dosyalar[0])
        if self.parcalar:
            nolar = ", ".join(str(p.no) for p in sorted(self.parcalar, key=lambda p: p.no))
            return "%s  [bölüm %s]" % (ad, nolar)
        return ad


def dogal_anahtar(metin: str):
    """'dosya 2' < 'dosya 10' olacak şekilde sayıları sayı olarak sıralayan anahtar."""
    return [(0, int(p), "") if p.isdigit() else (1, 0, p) for p in re.split(r"(\d+)", metin) if p]
