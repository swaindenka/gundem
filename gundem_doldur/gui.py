"""Gündem Doldurucu - masaüstü arayüzü (tkinter)."""
from __future__ import annotations

import datetime as dt
import os
import queue
import sys
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import List, Optional

from . import __version__
from .bolumler import Dava, Parca, bolum_bilgisi
from .extract import CaseData, extract_case, looks_complete, norm
from .reader import read_document
from .writer import AYLAR, build_gundem, gundem_basligi

# (arayüzde görünen ad, writer.build_gundem 'duzen' değeri)
DUZEN_SECENEKLERI = [
    ("Otomatik: her dava tek sayfaya sığsın", "otomatik"),
    ("Klasik: şablondaki gibi", "klasik"),
    ("Kompakt: heyet başlığı solda", "kompakt"),
    ("Kompakt, küçük yazı (10 pt)", "kompakt-10"),
    ("Kompakt, daha küçük yazı (9 pt)", "kompakt-9"),
    ("Kompakt, en küçük yazı (8 pt)", "kompakt-8"),
]

ALANLAR = [
    # (anahtar, etiket, satır sayısı; 1 = tek satırlık giriş)
    ("sira_no", "Sıra No", 1),
    ("esas", "Esas No", 1),
    ("mahkeme", "Mahkeme", 1),
    ("daire", "Daire", 1),
    ("daire_karar", "Daire Kararı", 2),
    ("konu", "Konusu", 2),
    ("alt_bolum", "Hakimler\n(******* altı)", 12),
]


def resource_path(*parts: str) -> str:
    base = getattr(sys, "_MEIPASS", None) or os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, *parts)


def sonraki_carsamba(bugun: Optional[dt.date] = None) -> dt.date:
    bugun = bugun or dt.date.today()
    fark = (2 - bugun.weekday()) % 7
    return bugun + dt.timedelta(days=fark or 7)


class App(tk.Tk):
    def __init__(self, ilk_dosyalar: List[str]):
        super().__init__()
        self.title("Gündem Doldurucu %s" % __version__)
        self.geometry("1280x740")
        self.minsize(900, 600)
        self.davalar: List[Dava] = []
        self.current: Optional[int] = None
        self.template = resource_path("sablon", "BASLIK_TEST.docx")
        self.q: "queue.Queue" = queue.Queue()
        self._build()
        self.after(100, self._poll)
        if ilk_dosyalar:
            self._load_files(ilk_dosyalar)

    @property
    def cases(self) -> List[CaseData]:
        return [d.veri for d in self.davalar]

    # ------------------------------------------------------------------ UI
    def _build(self):
        style = ttk.Style(self)
        try:
            style.theme_use("vista" if sys.platform == "win32" else "clam")
        except tk.TclError:
            pass

        top = ttk.Frame(self, padding=8)
        top.pack(fill="x")

        tarih_kutu = ttk.LabelFrame(top, text="Gündem tarihi", padding=6)
        tarih_kutu.pack(side="left")
        varsayilan = sonraki_carsamba()
        self.gun = tk.StringVar(value=str(varsayilan.day))
        self.ay = tk.StringVar(value=AYLAR[varsayilan.month - 1])
        self.yil = tk.StringVar(value=str(varsayilan.year))
        ttk.Spinbox(tarih_kutu, from_=1, to=31, width=4, textvariable=self.gun,
                    command=self._tarih_degisti).pack(side="left")
        ay_cb = ttk.Combobox(tarih_kutu, values=AYLAR, width=10, state="readonly", textvariable=self.ay)
        ay_cb.pack(side="left", padx=4)
        ay_cb.bind("<<ComboboxSelected>>", lambda e: self._tarih_degisti())
        ttk.Spinbox(tarih_kutu, from_=2000, to=2100, width=6, textvariable=self.yil,
                    command=self._tarih_degisti).pack(side="left")
        for v in (self.gun, self.yil):
            v.trace_add("write", lambda *a: self._tarih_degisti())
        self.baslik_onizleme = ttk.Label(tarih_kutu, width=42, foreground="#1a4d8f")
        self.baslik_onizleme.pack(side="left", padx=8)

        hakim_kutu = ttk.LabelFrame(top, text="Tetkik hakimi (isteğe bağlı)", padding=6)
        hakim_kutu.pack(side="left", padx=8, fill="y")
        self.tetkik = tk.StringVar(value="")
        ttk.Entry(hakim_kutu, textvariable=self.tetkik, width=28).pack()

        duzen_kutu = ttk.LabelFrame(top, text="Sayfa düzeni", padding=6)
        duzen_kutu.pack(side="left", padx=8, fill="y")
        self.duzen = tk.StringVar(value=DUZEN_SECENEKLERI[0][0])
        ttk.Combobox(duzen_kutu, textvariable=self.duzen, state="readonly", width=34,
                     values=[e for e, _ in DUZEN_SECENEKLERI]).pack()

        main = ttk.PanedWindow(self, orient="horizontal")
        main.pack(fill="both", expand=True, padx=8)

        # Sol: dosya listesi
        left = ttk.Frame(main, padding=(0, 0, 6, 0))
        main.add(left, weight=1)
        ttk.Label(left, text="Kaynak dosyalar (gündem sırası)").pack(anchor="w")
        self.liste = tk.Listbox(left, activestyle="none", exportselection=False)
        self.liste.pack(fill="both", expand=True, pady=4)
        self.liste.bind("<<ListboxSelect>>", self._secim_degisti)
        btns = ttk.Frame(left)
        btns.pack(fill="x")
        ttk.Button(btns, text="Dosya Ekle...", command=self._dosya_ekle).pack(side="left")
        ttk.Button(btns, text="Kaldır", command=self._kaldir).pack(side="left", padx=4)
        ttk.Button(btns, text="▲", width=3, command=lambda: self._tasi(-1)).pack(side="left")
        ttk.Button(btns, text="▼", width=3, command=lambda: self._tasi(1)).pack(side="left", padx=4)

        # Sağ: seçili dosyanın alanları
        right = ttk.Frame(main)
        main.add(right, weight=3)
        self.dosya_etiket = ttk.Label(right, text="Dosya ekleyin: .docx veya .doc", font=("", 10, "bold"))
        self.dosya_etiket.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 4))
        self.widgets = {}
        for i, (key, label, lines) in enumerate(ALANLAR, start=1):
            ttk.Label(right, text=label).grid(row=i, column=0, sticky="nw", padx=(0, 6), pady=2)
            if lines == 1:
                w = ttk.Entry(right)
            else:
                w = tk.Text(right, height=lines, wrap="word", undo=True, font=("Times New Roman", 11))
            w.grid(row=i, column=1, sticky="nsew", pady=2)
            self.widgets[key] = w
        right.columnconfigure(1, weight=1)
        right.rowconfigure(len(ALANLAR), weight=1)
        self.uyari = ttk.Label(right, text="", foreground="#b00020", wraplength=700, justify="left")
        self.uyari.grid(row=len(ALANLAR) + 1, column=0, columnspan=2, sticky="w", pady=4)
        ttk.Label(
            right, foreground="#555", wraplength=700, justify="left",
            text="Hakimler kutusunda her satır tabloda ayrı bir satır olur: 'BAŞKAN/ÜYE :' ile başlayanlar "
                 "değerlendirme kutulu hakim satırı, ':' ile bitenler ('... hakimleri:') kalın başlık satırı, "
                 "diğerleri (ilk derece hakimi gibi) değerlendirme kutulu tek satır olarak yazılır.",
        ).grid(row=len(ALANLAR) + 2, column=0, columnspan=2, sticky="w")

        bottom = ttk.Frame(self, padding=8)
        bottom.pack(fill="x")
        self.durum = ttk.Label(bottom, text="Hazır.")
        self.durum.pack(side="left")
        ttk.Button(bottom, text="Gündemi Oluştur (.docx)", command=self._olustur).pack(side="right")
        ttk.Button(bottom, text="Şablon Seç...", command=self._sablon_sec).pack(side="right", padx=6)
        self.sablon_etiket = ttk.Label(bottom, text="Şablon: yerleşik", foreground="#555")
        self.sablon_etiket.pack(side="right", padx=6)
        self._tarih_degisti()

    # ------------------------------------------------------------ Tarih
    def _tarih(self) -> Optional[dt.date]:
        try:
            return dt.date(int(self.yil.get()), AYLAR.index(self.ay.get()) + 1, int(self.gun.get()))
        except (ValueError, TypeError):
            return None

    def _tarih_degisti(self):
        t = self._tarih()
        self.baslik_onizleme.config(text=gundem_basligi(t) if t else "Geçersiz tarih")

    # ------------------------------------------------------------ Dosyalar
    def _dosya_ekle(self):
        paths = filedialog.askopenfilenames(
            title="Kaynak Word belgelerini seçin",
            filetypes=[("Word belgeleri", "*.docx *.doc"), ("Tüm dosyalar", "*.*")],
        )
        if paths:
            self._load_files(list(paths))

    def _load_files(self, paths: List[str]):
        self.durum.config(text="Okunuyor: %d dosya..." % len(paths))
        self.config(cursor="watch")

        def sirala(p):
            grup, no = bolum_bilgisi(p)
            ad = grup or norm(os.path.splitext(os.path.basename(p))[0]).lower()
            return (ad, no or 0)

        def work():
            for p in sorted(paths, key=sirala):
                logs: List[str] = []
                try:
                    grup, no = bolum_bilgisi(p)
                    doc = read_document(p, log=logs.append,
                                        is_good=lambda d: looks_complete(d, bolum=no is not None))
                    case = extract_case(doc, kaynak=p)
                    self.q.put(("case", (p, grup, no, case, logs)))
                except Exception as e:
                    traceback.print_exc()
                    self.q.put(("error", "%s\n\n%s" % (os.path.basename(p), e)))
            self.q.put(("done", None))

        threading.Thread(target=work, daemon=True).start()

    def _poll(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == "case":
                    self._save_current()
                    self._dava_ekle(*payload)
                elif kind == "error":
                    messagebox.showerror("Dosya okunamadı", payload)
                elif kind == "done":
                    self.config(cursor="")
                    self.durum.config(text="%d dava yüklü." % len(self.davalar))
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def _dava_ekle(self, yol: str, grup: str, no: Optional[int], case: CaseData, logs: List[str]):
        if no is not None:
            parca = Parca(no=no, yol=yol, veri=case, kayit=logs)
            mevcut = next((i for i, d in enumerate(self.davalar) if d.grup and d.grup == grup), None)
            if mevcut is not None:
                # Aynı davanın başka bir bölümü: mevcut davayla birleştir
                self.davalar[mevcut].parca_ekle(parca)
                self.liste.delete(mevcut)
                self.liste.insert(mevcut, self._liste_metni(self.davalar[mevcut]))
                self.current = None
                self._sec(mevcut)
                return
            dava = Dava.bolumden(grup, parca)
        else:
            case.uyarilar = logs + case.uyarilar
            dava = Dava.tek_dosya(case)
        if not dava.veri.sira_no:
            nums = [int(c.sira_no) for c in self.cases if c.sira_no.strip().isdigit()]
            dava.veri.sira_no = str(max(nums) + 1 if nums else len(self.davalar) + 1)
        self.davalar.append(dava)
        self.liste.insert("end", self._liste_metni(dava))
        self._sec(len(self.davalar) - 1)

    @staticmethod
    def _liste_metni(dava: Dava) -> str:
        return "%s. %s" % (dava.veri.sira_no or "?", dava.etiket())

    def _secim_degisti(self, _event=None):
        sel = self.liste.curselection()
        if not sel:
            return
        idx = sel[0]
        self._save_current()
        self._sec(idx)

    def _sec(self, idx: int):
        self.liste.selection_clear(0, "end")
        self.liste.selection_set(idx)
        self.liste.see(idx)
        self._show(idx)

    def _show(self, idx: int):
        self.current = idx
        dava = self.davalar[idx]
        case = dava.veri
        self.dosya_etiket.config(text="  +  ".join(os.path.basename(p) for p in dava.dosyalar))
        for key, _, lines in ALANLAR:
            w = self.widgets[key]
            val = getattr(case, key)
            if lines == 1:
                w.delete(0, "end")
                w.insert(0, val)
            else:
                w.delete("1.0", "end")
                w.insert("1.0", val)
        self.uyari.config(text="\n".join(case.uyarilar))

    def _save_current(self):
        if self.current is None or self.current >= len(self.davalar):
            return
        case = self.davalar[self.current].veri
        for key, _, lines in ALANLAR:
            w = self.widgets[key]
            val = w.get() if lines == 1 else w.get("1.0", "end-1c")
            setattr(case, key, val.strip("\n") if lines > 1 else val.strip())
        sel = self.liste.curselection()
        self.liste.delete(self.current)
        self.liste.insert(self.current, self._liste_metni(self.davalar[self.current]))
        for i in sel:
            self.liste.selection_set(i)

    def _clear_fields(self):
        self.current = None
        for key, _, lines in ALANLAR:
            w = self.widgets[key]
            if lines == 1:
                w.delete(0, "end")
            else:
                w.delete("1.0", "end")
        self.dosya_etiket.config(text="Dosya ekleyin: .docx veya .doc")
        self.uyari.config(text="")

    def _kaldir(self):
        sel = self.liste.curselection()
        if not sel:
            return
        idx = sel[0]
        self.current = None
        del self.davalar[idx]
        self.liste.delete(idx)
        if self.davalar:
            self._sec(min(idx, len(self.davalar) - 1))
        else:
            self._clear_fields()

    def _tasi(self, yon: int):
        sel = self.liste.curselection()
        if not sel:
            return
        self._save_current()
        i = sel[0]
        j = i + yon
        if not (0 <= j < len(self.davalar)):
            return
        a, b = self.davalar[i], self.davalar[j]
        # Sıra numaraları sırada kalsın: numaraları da değiş tokuş et
        a.veri.sira_no, b.veri.sira_no = b.veri.sira_no, a.veri.sira_no
        self.davalar[i], self.davalar[j] = b, a
        for k in (i, j):
            self.liste.delete(k)
            self.liste.insert(k, self._liste_metni(self.davalar[k]))
        self.current = None
        self._sec(j)

    # ------------------------------------------------------------ Çıktı
    def _sablon_sec(self):
        p = filedialog.askopenfilename(title="Gündem şablonunu seçin", filetypes=[("Word belgesi", "*.docx")])
        if p:
            self.template = p
            self.sablon_etiket.config(text="Şablon: %s" % os.path.basename(p))

    def _olustur(self):
        self._save_current()
        if not self.cases:
            messagebox.showwarning("Dosya yok", "Önce en az bir kaynak dosya ekleyin.")
            return
        tarih = self._tarih()
        if tarih is None:
            messagebox.showwarning("Tarih", "Geçerli bir gündem tarihi seçin.")
            return
        varsayilan = "GUNDEM_%s.docx" % tarih.strftime("%Y-%m-%d")
        ilk_klasor = os.path.dirname(self.cases[0].kaynak)
        out = filedialog.asksaveasfilename(
            title="Gündemi kaydet", defaultextension=".docx", initialfile=varsayilan,
            initialdir=ilk_klasor, filetypes=[("Word belgesi", "*.docx")],
        )
        if not out:
            return
        duzen = dict(DUZEN_SECENEKLERI).get(self.duzen.get(), "otomatik")
        try:
            secimler = build_gundem(self.template, self.cases, tarih, out, self.tetkik.get(), duzen)
        except PermissionError:
            messagebox.showerror("Kaydedilemedi", "Dosya başka bir programda açık olabilir:\n%s" % out)
            return
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Hata", "Gündem oluşturulamadı:\n%s" % e)
            return
        self.durum.config(text="Kaydedildi: %s" % out)
        notlar = []
        sigmayan = [c.sira_no or str(i + 1) for i, (c, (_, sigar)) in enumerate(zip(self.cases, secimler)) if not sigar]
        sikistirilan = [c.sira_no or str(i + 1) for i, (c, (ad, _)) in enumerate(zip(self.cases, secimler))
                        if ad != "klasik"]
        if duzen == "otomatik" and sikistirilan:
            notlar.append("Tek sayfaya sığması için kompakt düzen / küçük yazı kullanılan sıra no: %s"
                          % ", ".join(sikistirilan))
        if sigmayan:
            notlar.append("DİKKAT: Şu sıra no'lardaki davalar çok sayıda hakim nedeniyle tek sayfaya "
                          "sığmayabilir: %s" % ", ".join(sigmayan))
        ek = ("\n\n" + "\n".join(notlar)) if notlar else ""
        if messagebox.askyesno("Tamamlandı", "Gündem oluşturuldu:\n%s%s\n\nŞimdi açılsın mı?" % (out, ek)):
            try:
                if sys.platform == "win32":
                    os.startfile(out)  # type: ignore[attr-defined]
                else:
                    import subprocess
                    subprocess.Popen(["xdg-open", out])
            except Exception:
                pass


def _self_test(out_path: str) -> None:
    """Derlenmiş .exe'nin çalıştığını CI'da doğrulamak için: şablonu ve arayüzü yükleyip çıkar."""
    from .writer import _Template
    _Template(resource_path("sablon", "BASLIK_TEST.docx"))
    root = tk.Tk()
    root.update()
    root.destroy()
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("OK %s" % __version__)


def main():
    if len(sys.argv) == 3 and sys.argv[1] == "--self-test":
        _self_test(sys.argv[2])
        return
    dosyalar = [a for a in sys.argv[1:] if os.path.isfile(a)]
    App(dosyalar).mainloop()


if __name__ == "__main__":
    main()
