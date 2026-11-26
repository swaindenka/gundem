"""Belgelerin okunduktan sonra temsil edildiği basit veri modeli."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Callable, List, Optional, Union


@dataclass
class Cell:
    text: str
    span: int = 1           # kaç ızgara sütunu kapladığı (yatay birleştirme)
    vcontinue: bool = False  # dikey birleştirmenin devamı (üstteki hücre ile aynı)


@dataclass
class Table:
    rows: List[List[Cell]] = field(default_factory=list)

    def lines(self) -> List[str]:
        """Tablonun metnini satır satır döndürür (hücre içi satırlar dahil)."""
        out: List[str] = []
        for row in self.rows:
            for cell in row:
                if cell.vcontinue:
                    continue
                out.extend(cell.text.split("\n"))
        return out

    def row_lines(self, atla: Optional[Callable[[str], bool]] = None) -> List[str]:
        """Her tablo satırını tek metin satırı olarak döndürür (hücreler sekmeyle ayrılır).

        Hakim listeleri tablo içinde olabilir: "BAŞKAN | : | Ad Soyad | 12345". Hücreler ayrı
        satırlara bölünürse görev, isim ve sicil birbirinden kopar; bu yüzden satır korunur.
        Bir satırda birden çok hücre aynı sayıda alt satır içeriyorsa (görevler bir sütunda,
        isimler diğerinde) alt satırlar karşılıklı eşleştirilir.
        `atla`: True döndürdüğü hücreler (ör. değerlendirme kutuları) yok sayılır.
        """
        out: List[str] = []
        for row in self.rows:
            parts = []
            for cell in row:
                if cell.vcontinue:
                    continue
                lines = [l.strip() for l in cell.text.split("\n") if l.strip()]
                if not lines or (atla is not None and atla(cell.text)):
                    continue
                parts.append(lines)
            if not parts:
                continue
            multi = [p for p in parts if len(p) > 1]
            if len(multi) >= 2 and len({len(p) for p in multi}) == 1:
                k = len(multi[0])
                for i in range(k):
                    out.append("\t".join(p[i] if len(p) == k else (p[0] if i == 0 else "") for p in parts)
                               .strip("\t"))
            elif len(parts) == 1:
                out.extend(parts[0])
            else:
                out.append("\t".join(" ".join(p) for p in parts))
        return out


Block = Union[str, Table]  # str = paragraf, Table = tablo


@dataclass
class Document:
    blocks: List[Block] = field(default_factory=list)

    @property
    def tables(self) -> List[Table]:
        return [b for b in self.blocks if isinstance(b, Table)]

    def lines(self, satir_bazli: bool = False,
              atla: Optional[Callable[[str], bool]] = None) -> List[str]:
        """Belgenin tüm metnini sırayla, satır satır döndürür.

        satir_bazli=True: tablolarda her satır tek metin satırı olur (bkz. Table.row_lines).
        """
        out: List[str] = []
        for b in self.blocks:
            if isinstance(b, Table):
                out.extend(b.row_lines(atla) if satir_bazli else b.lines())
            else:
                out.extend(b.split("\n"))
        return out


# Yıldız benzeri karakterler: '*' dışında Word'de "Simge Ekle" ile ya da farklı yazı
# tiplerinden gelebilen yıldız / asterisk işaretleri.
YILDIZLAR = "*∗＊⁎✱✲✳❋٭﹡★☆⋆✶✴✵✷✸✹✺✻✼✽❉❊✦✧"


def sembol_karakteri(kod: int) -> str:
    """Word 'Symbol' / 'Wingdings' simge kodunu (ör. F02A) metne çevirir.

    Asterisk (0x2A) '*' olur; Wingdings yıldız aralığı '★' olur; diğerleri nesne işareti.
    """
    alt = kod & 0xFF
    if alt == 0x2A:
        return "*"
    if 0xA9 <= alt <= 0xB5:
        return "★"
    return "￼"
