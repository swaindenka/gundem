"""Belgelerin okunduktan sonra temsil edildiği basit veri modeli."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Union


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


Block = Union[str, Table]  # str = paragraf, Table = tablo


@dataclass
class Document:
    blocks: List[Block] = field(default_factory=list)

    @property
    def tables(self) -> List[Table]:
        return [b for b in self.blocks if isinstance(b, Table)]

    def lines(self) -> List[str]:
        """Belgenin tüm metnini sırayla, satır satır döndürür."""
        out: List[str] = []
        for b in self.blocks:
            if isinstance(b, Table):
                out.extend(b.lines())
            else:
                out.extend(b.split("\n"))
        return out
