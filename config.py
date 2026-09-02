"""Merkezi ayarlar: yollar, sabitler, arama kalıpları."""
from __future__ import annotations

from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent

DATA_DIR = BASE_DIR / "data"
INPUT_DIR = DATA_DIR / "input"
CACHE_DIR = DATA_DIR / "cache"
DOWNLOADS_DIR = DATA_DIR / "downloads"
OUTPUT_DIR = DATA_DIR / "output"

# Girdi referans dosyaları (kullanıcı tarafından sağlandı)
BIST100_LISTE_XLSX = INPUT_DIR / "bist100_liste.xlsx"
SIRKETLER_REFERANS_XLSX = INPUT_DIR / "sirketler_referans.xlsx"
TSRS_ENVANTERI_XLSX = INPUT_DIR / "tsrs_rapor_envanteri.xlsx"
MANUAL_IR_URLS_JSON = INPUT_DIR / "manual_ir_urls.json"

# Çalışma zamanı durumu
PROGRESS_FILE = DATA_DIR / "progress.json"
IR_URL_CACHE_FILE = CACHE_DIR / "company_websites.json"
REPORT_MANIFEST_FILE = CACHE_DIR / "report_manifest.json"
FINDINGS_CACHE_FILE = CACHE_DIR / "findings.json"

# Çıktı
OUTPUT_XLSX = OUTPUT_DIR / "kapsam3_bulgular.xlsx"

# Örneklem: sabit, 31 Aralık 2025 itibarıyla BIST 100
SCAN_YEARS = [2023, 2024, 2025]
CUTOFF_DATE = "2026-09-01"  # bu tarihe kadar yayımlanmış her versiyon dahil

REPORT_TYPES = [
    "Sürdürülebilirlik Raporu",
    "Faaliyet Raporu",
]

# Şirket başına parça (chunk) boyutu — limit aşımlarını önlemek için
CHUNK_SIZE = 15

# --- HTTP ayarları (downloader.py) ---
HTTP_USER_AGENT = (
    "Mozilla/5.0 (compatible; Scope3ResearchBot/1.0; +academic-research)"
)
HTTP_TIMEOUT_SECONDS = 30
HTTP_MAX_RETRIES = 3
HTTP_RETRY_BACKOFF_SECONDS = 2.0
HTTP_MIN_PDF_BYTES = 2048  # bundan küçük indirilen dosya "bozuk" sayılır

# IR sitesi taraması: ana sayfadan başlayıp rapor/yatırımcı ilişkileri ile
# alakalı görünen bağlantıları izleyerek en fazla kaç sayfa ziyaret edilsin
# (best-first). Şirket başına tek seferlik maliyet olduğu için (bkz.
# downloader.build_ir_page_refs) nispeten cömert tutulabilir.
IR_CRAWL_MAX_PAGES = 8

# KGK Rapor Envanteri (config.TSRS_ENVANTERI_XLSX) kaynaklı indirmeler:
# KGK, "Önerilen BIST Kodu" ile önceden başlık-eşleştirmesi yapılmış gerçek
# TSRS raporlarını içerdiği için IR sitesi taramasına göre ÖNCELİKLİ/daha
# güvenilir kaynaktır (bkz. build_ir_page_refs'in existing_keys ile KGK'da
# zaten olan kombinasyonları atlaması). Bu ortamdan kgk.gov.tr'ye giden
# isteklerin çoğu 500 Internal Server Error ile sonuçlanıyor (WAF/IP engeli
# olabilir) - kaynağı tamamen KAPATMAK yerine (bu, gerçek TSRS raporlarını
# kaçırmaya sebep olur) KGK_MAX_RETRIES/KGK_TIMEOUT_SECONDS ile hızlı-
# başarısızlık politikası uygulanır: tek deneme + kısa zaman aşımı, böylece
# hem WAF engeli geçiciyse yakalanır hem de başarısız denemeler 3×30sn yerine
# yalnızca ~8sn sürer.
ENABLE_KGK_SOURCE = True
KGK_MAX_RETRIES = 1
KGK_TIMEOUT_SECONDS = 8.0

# --- text_miner.py: anahtar kelime / desen tanımları ---
# Sayfa bağlamı alıntı uzunluğu (karakter, eşleşmenin etrafında yaklaşık ortalanmış)
CONTEXT_CHARS = 300

SCOPE3_KEYWORDS = [
    "kapsam 3",
    "scope 3",
    "değer zinciri",
    "value chain",
]

# GHG Protokolü Kapsam 3 kategorileri 1-15 (Kategori N / Category N)
CATEGORY_PATTERN = r"kategori\s*0?(1[0-5]|[1-9])\b|category\s*0?(1[0-5]|[1-9])\b"

ASSURANCE_KEYWORDS = [
    "güvence",
    "assurance",
    "iso 14064",
    "sınırlı güvence",
    "makul güvence",
    "bağımsız güvence",
]

NUMERIC_UNIT_PATTERNS = [
    r"tco2e",
    r"tco2-e",
    r"tco₂e",
    r"ton\s*co2\s*e",
    r"ton\s*co2e",
    r"co2\s*eşdeğeri",
    r"co2e\s*eşdeğeri",
]

MATURITY_REVIEW_NOTE = "ön etiket, manuel teyit gerekir"
