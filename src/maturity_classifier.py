"""Şeffaf, kural tabanlı (rule-based) olgunluk sınıflandırıcı.

Kurallar (öncelik sırasıyla — bir bulgu grubunda birden fazla sinyal varsa
EN YÜKSEK olgunluk seviyesi geçerli sayılır):
  1. Metodolojili-Güvenceli : "güvence" / "assurance" / "ISO 14064" geçiyorsa
  2. Kategori Bazlı          : "Kategori N" (N=1..15) / "Category N" geçiyorsa
  3. Toplam Sayısal          : "tCO2e" gibi bir emisyon birimi geçiyorsa
  4. Nitel                   : yalnızca "Kapsam 3"/"Scope 3"/"Değer Zinciri" geçiyorsa
  5. Açıklama Yok            : hiçbir sinyal yoksa (rapor indirildi ama eşleşme yok)

ÖNEMLİ: Bu sınıflandırma tamamen sezgiseldir ve otomatik üretilir. Her sonuç
`MaturityResult.manuel_teyit_notu` ile AÇIKÇA "ön etiket, manuel teyit
gerekir" olarak işaretlenir — akademik yayında nihai karar olarak
kullanılmadan önce elle doğrulanmalıdır.
"""
from __future__ import annotations

import re
from collections import defaultdict

import config
from src.models import Finding, MaturityResult, OlgunlukEtiketi

# Öncelik sırası: yüksekten düşüğe (classify_group bu sırayı kullanır)
_PRIORITY = [
    OlgunlukEtiketi.METODOLOJILI_GUVENCELI,
    OlgunlukEtiketi.KATEGORI_BAZLI,
    OlgunlukEtiketi.TOPLAM_SAYISAL,
    OlgunlukEtiketi.NITEL,
]

_NUMERIC_UNIT_RE = re.compile("|".join(config.NUMERIC_UNIT_PATTERNS), re.IGNORECASE)


def tag_finding(finding: Finding) -> OlgunlukEtiketi:
    """Tek bir ham bulguyu hangi olgunluk sinyaline ait olduğuna göre etiketler.

    Ham Bulgular sekmesindeki her satır için kullanılır (bkz. excel_writer.py).
    """
    terim_low = finding.eslesen_terim.lower()

    if any(kw in terim_low for kw in config.ASSURANCE_KEYWORDS):
        return OlgunlukEtiketi.METODOLOJILI_GUVENCELI
    if finding.kategori_no is not None:
        return OlgunlukEtiketi.KATEGORI_BAZLI
    if _NUMERIC_UNIT_RE.search(terim_low):
        return OlgunlukEtiketi.TOPLAM_SAYISAL
    return OlgunlukEtiketi.NITEL


def classify_group(kod: str, yil: int, rapor_turu: str, findings: list[Finding]) -> MaturityResult:
    """Bir şirket-yıl-rapor türü grubundaki tüm bulguları tek bir olgunluk sonucuna indirger."""
    if not findings:
        return MaturityResult(
            kod=kod,
            yil=yil,
            rapor_turu=rapor_turu,
            etiket=OlgunlukEtiketi.ACIKLAMA_YOK,
            gerekce="Rapor tarandı, Kapsam 3 ile ilgili hiçbir terim bulunamadı.",
        )

    tags = {tag_finding(f) for f in findings}
    etiket = next((t for t in _PRIORITY if t in tags), OlgunlukEtiketi.NITEL)

    kategori_kapsami = sorted({f.kategori_no for f in findings if f.kategori_no is not None})

    gerekce_parcalari = []
    for tag in _PRIORITY:
        if tag in tags:
            count = sum(1 for f in findings if tag_finding(f) == tag)
            gerekce_parcalari.append(f"{tag.value} sinyali ({count} eşleşme)")
    gerekce = "; ".join(gerekce_parcalari)

    return MaturityResult(
        kod=kod,
        yil=yil,
        rapor_turu=rapor_turu,
        etiket=etiket,
        kategori_kapsami=kategori_kapsami,
        gerekce=gerekce,
    )


def classify_all(findings: list[Finding]) -> list[MaturityResult]:
    """Tüm bulguları (kod, yıl, rapor_turu) bazında gruplayıp sınıflandırır.

    NOT: Yalnızca en az bir bulgusu olan gruplar için sonuç üretir. "Rapor
    indirildi ama hiç bulgu yok" (Açıklama Yok) ve "Rapor Yok" durumları
    main.py'de ReportRef manifest'i ile birleştirilerek Özet sekmesine eklenir
    (bkz. excel_writer.build_summary_rows).
    """
    grouped: dict[tuple[str, int, str], list[Finding]] = defaultdict(list)
    for f in findings:
        grouped[(f.kod, f.yil, f.rapor_turu)].append(f)

    return [classify_group(kod, yil, rapor_turu, group) for (kod, yil, rapor_turu), group in grouped.items()]
