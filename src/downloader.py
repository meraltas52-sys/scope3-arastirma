"""Her şirket-yıl-rapor türü için PDF indirir. İdempotent: zaten indirilmiş
ve geçerli bir dosya varsa tekrar indirmez.

Rapor referansları iki kaynaktan derlenir:
  1. KGK Rapor Envanteri (config.TSRS_ENVANTERI_XLSX) — önceden derlenmiş,
     "Önerilen BIST Kodu" ile şirket eşleştirmesi yapılmış 2023-2025
     sürdürülebilirlik/entegre rapor bağlantıları.
  2. Şirketin kendi yatırımcı ilişkileri sayfası (inventory.py cache'i) —
     KGK'da bulunmayan / faaliyet raporu gibi ek belgeler için sayfa
     taranarak PDF bağlantıları aranır.

Her iki kaynak için de bulunamayan şirket-yıl-rapor türü kombinasyonu
"Rapor Yok" olarak manifest'e yazılır (excel_writer.py "Eksik Raporlar"
sekmesinde kullanır) — atlanmaz, işaretlenir.

NOT (ağ erişimi): Bu modüldeki tüm indirme/tarama işlemleri canlı HTTP
istekleri gerektirir. Dış ağa kapalı bir ortamda çalıştırıldığında her
istek zaman aşımına uğrar/başarısız olur; her başarısızlık o rapor için
"Hata" ya da "Rapor Yok" olarak işaretlenir, akış durmaz.
"""
from __future__ import annotations

import logging
import re
import time
from dataclasses import asdict
from pathlib import Path
from urllib.parse import urljoin, urlparse

import openpyxl
import requests

import config
from src.inventory import load_cache as load_ir_cache
from src.models import Company, InvestorRelationsInfo, ReportRef, ReportStatus
from src.progress import ProgressStore, load_json, run_in_chunks, save_json

logger = logging.getLogger(__name__)

try:
    from bs4 import BeautifulSoup
except ImportError:  # pragma: no cover - requirements.txt'de var, savunma amaçlı
    BeautifulSoup = None


# Tarayıcı gibi görünen istek başlıkları — bazı sunucular (örn. KGK) sade
# "User-Agent" içeren istekleri bot olarak algılayıp 500/403 dönebiliyor.
BROWSER_HEADERS = {
    "User-Agent": config.HTTP_USER_AGENT,
    "Accept": "application/pdf,application/octet-stream,text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "tr-TR,tr;q=0.9,en-US;q=0.8,en;q=0.7",
    "Referer": "https://www.kgk.gov.tr/",
    "Connection": "keep-alive",
}

# İstekler arası minimum bekleme (saniye) — hızlı art arda istek atmak
# bazı sunucularda geçici engellemeye/500 hatasına yol açabiliyor.
REQUEST_DELAY_SECONDS = 1.5


# --------------------------------------------------------------------------
# Rapor türü sınıflandırma yardımcıları
# --------------------------------------------------------------------------

def infer_report_type(title: str) -> str:
    """Rapor başlığından "Sürdürülebilirlik Raporu" / "Faaliyet Raporu" ayrımı yapar.

    KGK listesi "Entegre Faaliyet Raporu" gibi birleşik başlıklar da içerir;
    bunlar sürdürülebilirlik içeriği taşıdığından "Sürdürülebilirlik Raporu"
    kovasına dahil edilir. Yalnızca "sürdürülebilirlik/entegre" geçmeyen saf
    "faaliyet raporu" başlıkları "Faaliyet Raporu" olarak sınıflanır.
    """
    t = title.lower()
    if "sürdürülebilirlik" in t or "entegre" in t or "surdurulebilirlik" in t:
        return "Sürdürülebilirlik Raporu"
    if "faaliyet" in t:
        return "Faaliyet Raporu"
    return "Sürdürülebilirlik Raporu"


# --------------------------------------------------------------------------
# Kaynak 1: KGK Rapor Envanteri
# --------------------------------------------------------------------------

def load_tsrs_envanteri_refs(companies_by_kod: dict[str, Company], path=config.TSRS_ENVANTERI_XLSX) -> list[ReportRef]:
    """Önceden derlenmiş KGK envanterini BIST100 örneklemiyle eşleştirip ReportRef üretir.

    BIST100 dışındaki eşleşmeler SESSİZCE ATLANMAZ: sadece bu fonksiyonun
    döndürdüğü listeye girmezler (örneklem dışı oldukları için); tam envanter
    karşılaştırması için bkz. `diff_tsrs_with_bist100`.
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Rapor Envanteri"]
    rows = list(ws.iter_rows(values_only=True))
    header = [str(c).strip() if c else "" for c in rows[0]]
    idx = {name: i for i, name in enumerate(header)}

    refs: list[ReportRef] = []
    for row in rows[1:]:
        yil = row[idx["Yıl"]]
        baslik = row[idx["Rapor Başlığı"]] or ""
        url = row[idx["Kaynak URL"]]
        onerilen_kod = (row[idx.get("Önerilen BIST Kodu", -1)] or "").strip().upper() if idx.get("Önerilen BIST Kodu") is not None else ""
               if not yil or not url or onerilen_kod not in companies_by_kod:
            continue
        # Sanity check: rapor başlığı ile şirket ünvanı arasında en az bir
        # anlamlı ortak kelime olmalı - yoksa "Önerilen BIST Kodu" sütunundaki
        # otomatik eşleştirme hatalı olabilir (örn. TTKOM -> Kuveyt Türk gibi).
        company_unvan = companies_by_kod[onerilen_kod].unvan.upper()
        baslik_upper = str(baslik).upper()
        company_words = {w for w in re.split(r"[^A-ZÇĞİÖŞÜ0-9]+", company_unvan) if len(w) > 3
                          and w not in {"A.Ş", "A.O", "SANAYİ", "TİCARET", "HOLDİNG", "VE"}}
        if company_words and not any(w in baslik_upper for w in company_words):
            logger.warning(
                "Şüpheli eşleşme atlandı: %s koduna önerilen rapor '%s' başlıkla uyuşmuyor.",
                onerilen_kod, baslik,
            )
            continue
        refs.append(
            ReportRef(
                kod=onerilen_kod,
                yil=int(yil),
                rapor_turu=infer_report_type(baslik),
                kaynak_url=url,
                kaynak_turu="KGK",
                durum=ReportStatus.INDIRILECEK,
            )
        )
    return refs


def diff_tsrs_with_bist100(companies_by_kod: dict[str, Company], path=config.TSRS_ENVANTERI_XLSX) -> dict:
    """TSRS envanterindeki 150 raporu BIST100 örneklemiyle karşılaştırır.

    Döner: {
      "bist100_disi_satirlar": [...],       # envanterde var ama BIST100'de yok
      "bist100_raporu_olmayan_sirketler": [...],  # BIST100'de var, envanterde HİÇ raporu yok
    }
    Envanteri SİLMEZ/değiştirmez — sadece raporlar.
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Rapor Envanteri"]
    rows = list(ws.iter_rows(values_only=True))
    header = [str(c).strip() if c else "" for c in rows[0]]
    idx = {name: i for i, name in enumerate(header)}

    bist100_disi = []
    kodlar_envanterde = set()
    for row in rows[1:]:
        onerilen_kod = (row[idx.get("Önerilen BIST Kodu", -1)] or "").strip().upper()
        baslik = row[idx["Rapor Başlığı"]]
        if onerilen_kod and onerilen_kod not in companies_by_kod:
            bist100_disi.append({"kod": onerilen_kod, "baslik": baslik})
        elif onerilen_kod:
            kodlar_envanterde.add(onerilen_kod)

    hic_raporu_olmayan = sorted(set(companies_by_kod) - kodlar_envanterde)
    return {
        "bist100_disi_satirlar": bist100_disi,
        "bist100_raporu_olmayan_sirketler": hic_raporu_olmayan,
    }


# --------------------------------------------------------------------------
# Kaynak 2: Şirket yatırımcı ilişkileri sayfası taraması
# --------------------------------------------------------------------------

_PDF_LINK_RE = re.compile(r'href=["\']([^"\']+\.pdf)["\']', re.IGNORECASE)


def discover_pdf_links_from_ir_page(
    ir_url: str, session: requests.Session, timeout: float = config.HTTP_TIMEOUT_SECONDS
) -> list[str]:
    """IR sayfasını çeker ve .pdf ile biten bağlantıları döner. Ağ kapalıysa [] döner."""
    try:
        time.sleep(REQUEST_DELAY_SECONDS)
        resp = session.get(ir_url, timeout=timeout, headers=BROWSER_HEADERS)
        resp.raise_for_status()
    except requests.RequestException as exc:
        logger.debug("IR sayfası alınamadı (%s): %s", ir_url, exc)
        return []

    links: set[str] = set()
    if BeautifulSoup is not None:
        soup = BeautifulSoup(resp.text, "html.parser")
        for a in soup.find_all("a", href=True):
            if a["href"].lower().endswith(".pdf"):
                links.add(urljoin(ir_url, a["href"]))
    else:  # yedek: basit regex
        for match in _PDF_LINK_RE.finditer(resp.text):
            links.add(urljoin(ir_url, match.group(1)))

    return sorted(links)


def pick_best_pdf_link(links: list[str], yil: int, rapor_turu: str) -> str | None:
    """Aday PDF bağlantıları arasından yıl + rapor türü ipuçlarına en uygun olanı seçer."""
    yil_str = str(yil)
    type_keywords = (
        ["surdurulebilirlik", "sustainability", "entegre", "integrated"]
        if rapor_turu == "Sürdürülebilirlik Raporu"
        else ["faaliyet", "annual"]
    )

    def normalize(s: str) -> str:
        return s.lower().translate(str.maketrans("ışğüöç", "isguoc"))

    scored = []
    for link in links:
        name = normalize(urlparse(link).path)
        score = (yil_str in name) + sum(kw in name for kw in type_keywords)
        if score > 0:
            scored.append((score, link))

    if not scored:
        return None
    scored.sort(key=lambda t: t[0], reverse=True)
    return scored[0][1]


def build_ir_page_refs(
    companies: list[Company],
    ir_cache: dict[str, InvestorRelationsInfo],
    existing_keys: set[tuple[str, int, str]],
    session: requests.Session,
) -> list[ReportRef]:
    """KGK envanterinde eksik kalan şirket-yıl-rapor türü kombinasyonları için IR sayfasını dener."""
    refs: list[ReportRef] = []
    for company in companies:
        ir_info = ir_cache.get(company.kod)
        ir_url = ir_info.yatirimci_iliskileri_url or ir_info.resmi_web_sitesi if ir_info else None

        for yil in config.SCAN_YEARS:
            for rapor_turu in config.REPORT_TYPES:
                key = (company.kod, yil, rapor_turu)
                if key in existing_keys:
                    continue
                if not ir_url:
                    refs.append(
                        ReportRef(kod=company.kod, yil=yil, rapor_turu=rapor_turu, durum=ReportStatus.BULUNMADI)
                    )
                    continue

                links = discover_pdf_links_from_ir_page(ir_url, session)
                best = pick_best_pdf_link(links, yil, rapor_turu)
                if best:
                    refs.append(
                        ReportRef(
                            kod=company.kod,
                            yil=yil,
                            rapor_turu=rapor_turu,
                            kaynak_url=best,
                            kaynak_turu="Şirket IR Sitesi",
                            durum=ReportStatus.INDIRILECEK,
                        )
                    )
                else:
                    refs.append(
                        ReportRef(kod=company.kod, yil=yil, rapor_turu=rapor_turu, durum=ReportStatus.BULUNMADI)
                    )
    return refs


# --------------------------------------------------------------------------
# İndirme (idempotent)
# --------------------------------------------------------------------------

def _slug(text: str) -> str:
    return re.sub(r"[^a-zA-Z0-9]+", "_", text).strip("_")


def local_path_for(kod: str, yil: int, rapor_turu: str) -> Path:
    return config.DOWNLOADS_DIR / kod / str(yil) / f"{_slug(rapor_turu)}.pdf"


def _is_valid_pdf(path: Path) -> bool:
    if not path.exists() or path.stat().st_size < config.HTTP_MIN_PDF_BYTES:
        return False
    with path.open("rb") as f:
        return f.read(5) == b"%PDF-"


def download_report(ref: ReportRef, session: requests.Session) -> ReportRef:
    """Tek bir raporu indirir. İdempotent: geçerli bir dosya zaten varsa atlar."""
    dest = local_path_for(ref.kod, ref.yil, ref.rapor_turu)

    if _is_valid_pdf(dest):
        ref.yerel_dosya = str(dest)
        ref.durum = ReportStatus.INDIRILDI
        return ref

    if not ref.kaynak_url:
        ref.durum = ReportStatus.BULUNMADI
        return ref

    dest.parent.mkdir(parents=True, exist_ok=True)
    last_exc: Exception | None = None
    for attempt in range(1, config.HTTP_MAX_RETRIES + 1):
        try:
            time.sleep(REQUEST_DELAY_SECONDS)
            resp = session.get(
                ref.kaynak_url,
                timeout=config.HTTP_TIMEOUT_SECONDS,
                headers=BROWSER_HEADERS,
                stream=True,
            )
            resp.raise_for_status()
            tmp_dest = dest.with_suffix(".part")
            with tmp_dest.open("wb") as f:
                for chunk in resp.iter_content(chunk_size=65536):
                    f.write(chunk)
            tmp_dest.replace(dest)

            if not _is_valid_pdf(dest):
                raise ValueError("İndirilen dosya geçerli bir PDF değil (magic bytes/boyut uyuşmuyor)")

            ref.yerel_dosya = str(dest)
            ref.durum = ReportStatus.INDIRILDI
            return ref
        except (requests.RequestException, ValueError, OSError) as exc:
            last_exc = exc
            logger.warning(
                "İndirme denemesi %d/%d başarısız (%s, %s, %s): %s",
                attempt, config.HTTP_MAX_RETRIES, ref.kod, ref.yil, ref.rapor_turu, exc,
            )
            if attempt < config.HTTP_MAX_RETRIES:
                time.sleep(config.HTTP_RETRY_BACKOFF_SECONDS * attempt)

    ref.durum = ReportStatus.HATA
    ref.hata_mesaji = str(last_exc)
    return ref


# --------------------------------------------------------------------------
# Manifest (text_miner.py / excel_writer.py bunu okur)
# --------------------------------------------------------------------------

def load_manifest(path=config.REPORT_MANIFEST_FILE) -> list[ReportRef]:
    raw = load_json(path, [])
    refs = []
    for item in raw:
        item = dict(item)
        item["durum"] = ReportStatus(item["durum"])
        refs.append(ReportRef(**item))
    return refs


def save_manifest(refs: list[ReportRef], path=config.REPORT_MANIFEST_FILE) -> None:
    save_json(path, [{**asdict(r), "durum": r.durum.value} for r in refs])


# --------------------------------------------------------------------------
# Parça bazlı çalıştırma
# --------------------------------------------------------------------------

def run(
    companies: list[Company],
    progress_store: ProgressStore,
    chunk_size: int = config.CHUNK_SIZE,
) -> list[ReportRef]:
    companies_by_kod = {c.kod: c for c in companies}
    ir_cache = load_ir_cache()
    session = requests.Session()

    tsrs_refs = load_tsrs_envanteri_refs(companies_by_kod)
    existing_keys = {r.key for r in tsrs_refs}
    ir_refs = build_ir_page_refs(companies, ir_cache, existing_keys, session)

    all_refs = tsrs_refs + ir_refs
    refs_by_key = {r.key: r for r in all_refs}

    def process(ref: ReportRef) -> None:
        updated = download_report(ref, session)
        refs_by_key[updated.key] = updated

    def on_chunk_done(chunk_idx: int, total_chunks: int, progress) -> None:
        save_manifest(list(refs_by_key.values()))
        indirildi = sum(1 for r in refs_by_key.values() if r.durum == ReportStatus.INDIRILDI)
        rapor_yok = sum(1 for r in refs_by_key.values() if r.durum == ReportStatus.BULUNMADI)
        hata = sum(1 for r in refs_by_key.values() if r.durum == ReportStatus.HATA)
        print(
            f"[downloader] Parça {chunk_idx}/{total_chunks} — "
            f"indirildi: {indirildi}, rapor yok: {rapor_yok}, hata: {hata}"
        )

    # Sadece indirilecek olanları (URL'si olanları) chunk'lı akışa sokuyoruz;
    # "Rapor Yok" olanlar zaten nihai durumda, boşuna işlenmez.
    to_download = [r for r in all_refs if r.kaynak_url]

    run_in_chunks(
        items=to_download,
        key_fn=lambda r: "|".join(map(str, r.key)),
        process_fn=process,
        progress_store=progress_store,
        stage_name="download",
        chunk_size=chunk_size,
        on_chunk_done=on_chunk_done,
    )

    save_manifest(list(refs_by_key.values()))
    return list(refs_by_key.values())


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    from src.inventory import load_bist100

    companies = load_bist100()
    progress_store = ProgressStore(config.PROGRESS_FILE)
    run(companies, progress_store)


if __name__ == "__main__":
    main()
