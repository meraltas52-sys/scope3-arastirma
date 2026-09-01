"""Pipeline boyunca paylaşılan veri modelleri."""
from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum


@dataclass
class Company:
    """BIST 100 örnekleminden bir şirket."""

    kod: str  # KAP/BIST kodu, örn. "AKBNK"
    unvan: str
    sehir: str | None = None
    bagimsiz_denetim: str | None = None  # Şirketler.xlsx referansından, varsa

    def __post_init__(self) -> None:
        self.kod = (self.kod or "").strip().upper()
        self.unvan = (self.unvan or "").strip()


@dataclass
class InvestorRelationsInfo:
    """inventory.py'nin şirket başına ürettiği/cache'lediği kayıt."""

    kod: str
    resmi_web_sitesi: str | None = None
    yatirimci_iliskileri_url: str | None = None
    kaynak: str | None = None  # "manuel", "otomatik-tahmin", "AI-on-bilgi-DOGRULANMADI" vb.
    dogrulandi: bool = False
    not_: str | None = None


class ReportStatus(str, Enum):
    BULUNMADI = "Rapor Yok"  # arandı, kaynak/URL yok
    INDIRILECEK = "İndirilecek"  # URL var, henüz indirilmedi
    INDIRILDI = "İndirildi"
    HATA = "Hata"


@dataclass
class ReportRef:
    """Bir şirket-yıl-rapor türü için tekil rapor referansı."""

    kod: str
    yil: int
    rapor_turu: str  # config.REPORT_TYPES
    kaynak_url: str | None = None
    kaynak_turu: str | None = None  # "KGK", "Şirket IR Sitesi"
    yerel_dosya: str | None = None
    durum: ReportStatus = ReportStatus.BULUNMADI
    hata_mesaji: str | None = None

    @property
    def key(self) -> tuple[str, int, str]:
        return (self.kod, self.yil, self.rapor_turu)


class OlgunlukEtiketi(str, Enum):
    ACIKLAMA_YOK = "Açıklama Yok"
    NITEL = "Nitel"
    TOPLAM_SAYISAL = "Toplam Sayısal"
    KATEGORI_BAZLI = "Kategori Bazlı"
    METODOLOJILI_GUVENCELI = "Metodolojili-Güvenceli"


@dataclass
class Finding:
    """text_miner.py'nin ürettiği ham bulgu satırı (denetlenebilirlik için)."""

    kod: str
    yil: int
    rapor_turu: str
    sayfa_no: int  # 1-indeksli
    eslesen_terim: str
    alinti: str  # ~300 karakterlik bağlam
    kategori_no: int | None = None  # Kategori 1-15 eşleşmesiyse


@dataclass
class MaturityResult:
    """maturity_classifier.py'nin şirket-yıl-rapor türü bazlı çıktısı."""

    kod: str
    yil: int
    rapor_turu: str
    etiket: OlgunlukEtiketi
    kategori_kapsami: list[int] = field(default_factory=list)  # bulunan Kategori N'ler
    gerekce: str = ""  # hangi sinyaller etikete yol açtı
    manuel_teyit_notu: str = "ön etiket, manuel teyit gerekir"
