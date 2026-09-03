"""PyMuPDF ile PDF'leri sayfa sayfa tarar; Kapsam 3 ile ilgili terimleri
bulur. Her eşleşme için sayfa numarası + ~300 karakterlik bağlam alıntısı
üretir (denetlenebilirlik: her bulgu kaynağına kadar izlenebilir).

Taranan terimler (config.py'de tanımlı):
  - Çekirdek Kapsam 3 terimleri: "Kapsam 3", "Scope 3", "Değer Zinciri", "Value Chain"
  - Kategori 1..15 / Category 1..15 (GHG Protokolü Kapsam 3 kategorileri)
  - Güvence/metodoloji terimleri (güvence, assurance, ISO 14064, ...)
  - Sayısal birim terimleri (tCO2e, ton CO2 eşdeğeri, ...)

Sezgisel kural: güvence/sayısal-birim terimleri yalnızca AYNI SAYFADA bir
çekirdek Kapsam 3 ya da Kategori N sinyali varsa kaydedilir — aksi halde
Kapsam 1/2 emisyonlarına ait alakasız eşleşmeler (gürültü) Ham Bulgular'a
karışabilir. Bu davranış maturity_classifier.py ile birlikte tasarlanmıştır.
"""
from __future__ import annotations

import logging
import re
from dataclasses import asdict
from pathlib import Path

import config
from src.models import Finding, ReportRef, ReportStatus
from src.progress import ProgressStore, load_json, run_in_chunks, save_json

logger = logging.getLogger(__name__)

try:
    import pymupdf as fitz  # PyMuPDF (eski adı `fitz`; yeni sürümlerde `pymupdf`)
except ImportError:  # pragma: no cover
    fitz = None

_CATEGORY_RE = re.compile(config.CATEGORY_PATTERN, re.IGNORECASE)


def _normalize_ws(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def _context_snippet(text: str, start: int, end: int, width: int = config.CONTEXT_CHARS) -> str:
    """Eşleşmenin etrafında ~width karakterlik, kabaca ortalanmış bağlam döner."""
    pad = max(width - (end - start), 0) // 2
    lo = max(0, start - pad)
    hi = min(len(text), end + pad)
    snippet = text[lo:hi]
    prefix = "…" if lo > 0 else ""
    suffix = "…" if hi < len(text) else ""
    return _normalize_ws(prefix + snippet + suffix)


def _find_keyword_findings(
    page_text: str, low_text: str, keywords: list[str], kod: str, yil: int, rapor_turu: str, sayfa_no: int
) -> list[Finding]:
    findings = []
    for kw in keywords:
        start = 0
        kw_low = kw.lower()
        while True:
            idx = low_text.find(kw_low, start)
            if idx == -1:
                break
            findings.append(
                Finding(
                    kod=kod,
                    yil=yil,
                    rapor_turu=rapor_turu,
                    sayfa_no=sayfa_no,
                    eslesen_terim=page_text[idx : idx + len(kw)],
                    alinti=_context_snippet(page_text, idx, idx + len(kw)),
                )
            )
            start = idx + len(kw)
    return findings


def _find_category_findings(
    page_text: str, low_text: str, kod: str, yil: int, rapor_turu: str, sayfa_no: int
) -> list[Finding]:
    findings = []
    for m in _CATEGORY_RE.finditer(low_text):
        kategori_no = int(m.group(1) or m.group(2))
        findings.append(
            Finding(
                kod=kod,
                yil=yil,
                rapor_turu=rapor_turu,
                sayfa_no=sayfa_no,
                eslesen_terim=page_text[m.start() : m.end()],
                alinti=_context_snippet(page_text, m.start(), m.end()),
                kategori_no=kategori_no,
            )
        )
    return findings


def _find_pattern_findings(
    page_text: str, low_text: str, patterns: list[str], kod: str, yil: int, rapor_turu: str, sayfa_no: int
) -> list[Finding]:
    findings = []
    for pattern in patterns:
        for m in re.finditer(pattern, low_text, re.IGNORECASE):
            findings.append(
                Finding(
                    kod=kod,
                    yil=yil,
                    rapor_turu=rapor_turu,
                    sayfa_no=sayfa_no,
                    eslesen_terim=page_text[m.start() : m.end()],
                    alinti=_context_snippet(page_text, m.start(), m.end()),
                )
            )
    return findings


def mine_pdf(path: Path, kod: str, yil: int, rapor_turu: str) -> list[Finding]:
    """Tek bir PDF'i sayfa sayfa tarar ve Finding listesi döner.

    Bozuk/okunamayan PDF durumunda TÜM İŞLEMİ DURDURMAZ: boş liste döner ve
    çağıran taraf (run()) bunu "hata" olarak işaretleyip devam eder.
    """
    if fitz is None:
        raise RuntimeError("PyMuPDF (pymupdf) kurulu değil: `pip install pymupdf`")

    findings: list[Finding] = []
    with fitz.open(path) as doc:
        for page_index, page in enumerate(doc):
            sayfa_no = page_index + 1  # 1-indeksli, kullanıcıya gösterilecek
            page_text = page.get_text()
            if not page_text.strip():
                continue
            # DİKKAT: düz .lower() KULLANMA - Python'un varsayılan Unicode
            # case-fold'u Türkçe büyük nokta'lı İ'yi (U+0130) TEK karakterden
            # İKİ karaktere ('i' + birleştirici nokta, U+0307) genişletir. Bu,
            # sayfadaki her İ'den sonra low_text'i page_text'e göre 1 karakter
            # kaydırıp aralarındaki hizalamayı bozuyordu - low_text üzerinde
            # bulunan eşleşme index'i (idx), page_text üzerinden yanlış yeri
            # dilimleyip "Eşleşen Terim"i (kısa dilim olduğu için tamamen
            # anlamsız hale gelecek şekilde) ve "Bağlam Alıntısı"nı (geniş
            # pencere olduğu için daha az fark edilir şekilde) bozuyordu. Önce
            # İ'yi TEK karakter 'i'ye çevirip hizalamayı koruyoruz.
            low_text = page_text.replace("İ", "i").lower()

            core = _find_keyword_findings(page_text, low_text, config.SCOPE3_KEYWORDS, kod, yil, rapor_turu, sayfa_no)
            kategori = _find_category_findings(page_text, low_text, kod, yil, rapor_turu, sayfa_no)

            page_findings = core + kategori
            if page_findings:
                # Sadece Kapsam 3 sinyali olan sayfalarda güvence/sayısal terimleri de ara.
                page_findings += _find_keyword_findings(
                    page_text, low_text, config.ASSURANCE_KEYWORDS, kod, yil, rapor_turu, sayfa_no
                )
                page_findings += _find_pattern_findings(
                    page_text, low_text, config.NUMERIC_UNIT_PATTERNS, kod, yil, rapor_turu, sayfa_no
                )

            findings.extend(page_findings)

    return findings


# --------------------------------------------------------------------------
# Bulgu önbelleği (excel_writer.py bunu okur)
# --------------------------------------------------------------------------

def load_findings(path=config.FINDINGS_CACHE_FILE) -> list[Finding]:
    raw = load_json(path, [])
    return [Finding(**item) for item in raw]


def save_findings(findings: list[Finding], path=config.FINDINGS_CACHE_FILE) -> None:
    save_json(path, [asdict(f) for f in findings])


# --------------------------------------------------------------------------
# Parça bazlı çalıştırma
# --------------------------------------------------------------------------

def run(manifest: list[ReportRef], progress_store: ProgressStore, chunk_size: int = config.CHUNK_SIZE) -> list[Finding]:
    indirilenler = [r for r in manifest if r.durum == ReportStatus.INDIRILDI and r.yerel_dosya]

    # Önceki çalıştırmanın bulgu önbelleğiyle başla: run_in_chunks,
    # progress.json'da zaten "tamamlandı" işaretli raporları process_fn'i
    # hiç çağırmadan atlar (kaldığı yerden devam). all_findings BOŞTAN
    # başlatılırsa, atlanan (önceden taranmış) raporların bulguları hiç
    # eklenmez ve dosya sonundaki save_findings çağrısı önbelleği SADECE bu
    # çalıştırmada gerçekten işlenen raporlarla değiştirir - önceki
    # çalıştırmalardaki tüm bulgular sessizce silinir (bkz. downloader.py
    # run()'daki aynı sınıftan hataya uygulanan previous_manifest_by_key
    # düzeltmesi).
    all_findings: dict[str, list[Finding]] = {}
    for finding in load_findings():
        key = "|".join(map(str, (finding.kod, finding.yil, finding.rapor_turu)))
        all_findings.setdefault(key, []).append(finding)

    def process(ref: ReportRef) -> None:
        key = "|".join(map(str, ref.key))
        all_findings[key] = mine_pdf(Path(ref.yerel_dosya), ref.kod, ref.yil, ref.rapor_turu)

    def on_chunk_done(chunk_idx: int, total_chunks: int, progress) -> None:
        flat = [f for findings in all_findings.values() for f in findings]
        save_findings(flat)
        print(
            f"[text_miner] Parça {chunk_idx}/{total_chunks} — "
            f"taranan rapor: {len(progress.completed)}, toplam bulgu: {len(flat)}, "
            f"hata: {len(progress.errors)}"
        )

    run_in_chunks(
        items=indirilenler,
        key_fn=lambda r: "|".join(map(str, r.key)),
        process_fn=process,
        progress_store=progress_store,
        stage_name="text_mining",
        chunk_size=chunk_size,
        on_chunk_done=on_chunk_done,
    )

    flat = [f for findings in all_findings.values() for f in findings]
    save_findings(flat)
    return flat


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    from src.downloader import load_manifest

    manifest = load_manifest()
    progress_store = ProgressStore(config.PROGRESS_FILE)
    run(manifest, progress_store)


if __name__ == "__main__":
    main()
