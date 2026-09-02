#!/usr/bin/env python3
"""BIST 100 Kapsam 3 (Scope 3) Emisyon Raporlama Olgunluğu Tarama Ajanı.

Tüm pipeline'ı sırayla çalıştırır:
  inventory -> downloader -> text_miner -> maturity_classifier -> excel_writer

Her aşama 15-20 şirketlik parçalar (chunk) halinde ilerler; ilerleme
data/progress.json içinde saklanır, tekrar çalıştırıldığında kaldığı yerden
devam eder. Bir şirkette/raporda hata oluşursa TÜM İŞLEM DURMAZ; hata
işaretlenip devam edilir, en sonda özet raporlanır.

Kullanım:
    python main.py                  # tüm aşamalar
    python main.py --stage inventory
    python main.py --stage download
    python main.py --stage mine
    python main.py --stage report
    python main.py --chunk-size 20
"""
from __future__ import annotations

import argparse
import logging

import requests

import config
from src.downloader import load_manifest, run as run_downloader
from src.excel_writer import write_workbook
from src.inventory import load_bist100, load_company_reference, run as run_inventory, search_company_website
from src.maturity_classifier import classify_all
from src.models import ReportStatus
from src.progress import ProgressStore
from src.text_miner import load_findings, run as run_text_miner

logger = logging.getLogger(__name__)

STAGES = ["inventory", "download", "mine", "report"]


def _print_banner(title: str) -> None:
    print(f"\n{'=' * 70}\n{title}\n{'=' * 70}")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument(
        "--stage", choices=STAGES, default=None,
        help="Sadece belirtilen aşamayı çalıştır (varsayılan: hepsi sırayla).",
    )
    parser.add_argument("--chunk-size", type=int, default=config.CHUNK_SIZE, help="Parça başına şirket sayısı.")
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

    companies = load_bist100()
    reference = load_company_reference()
    missing = [c.kod for c in companies if c.kod not in reference]
    if missing:
        logger.warning("788 şirketlik referansta bulunamayan BIST100 kodu: %s", ", ".join(missing))

    progress_store = ProgressStore(config.PROGRESS_FILE)
    stages_to_run = STAGES if args.stage is None else [args.stage]

    if "inventory" in stages_to_run:
        _print_banner(f"1/4 inventory — {len(companies)} şirket, parça boyutu {args.chunk_size}")
        _search_session = requests.Session()
        run_inventory(
            companies,
            progress_store,
            chunk_size=args.chunk_size,
            search_fn=lambda company: search_company_website(company, _search_session),
        )

    if "download" in stages_to_run:
        _print_banner("2/4 downloader — rapor indirme")
        manifest = run_downloader(companies, progress_store, chunk_size=args.chunk_size)
    else:
        manifest = load_manifest()

    if "mine" in stages_to_run:
        _print_banner("3/4 text_miner — PDF tarama")
        findings = run_text_miner(manifest, progress_store, chunk_size=args.chunk_size)
    else:
        findings = load_findings()

    if "report" in stages_to_run:
        _print_banner("4/4 maturity_classifier + excel_writer — rapor üretimi")
        maturity_results = classify_all(findings)
        write_workbook(config.OUTPUT_XLSX, companies, manifest, findings, maturity_results)
        print(f"Çıktı yazıldı: {config.OUTPUT_XLSX}")

    _print_summary(companies, manifest)


def _print_summary(companies, manifest) -> None:
    indirildi = sum(1 for r in manifest if r.durum == ReportStatus.INDIRILDI)
    rapor_yok = sum(1 for r in manifest if r.durum == ReportStatus.BULUNMADI)
    hata = sum(1 for r in manifest if r.durum == ReportStatus.HATA)
    indirilecek = sum(1 for r in manifest if r.durum == ReportStatus.INDIRILECEK)
    beklenen_toplam = len(companies) * len(config.SCAN_YEARS) * len(config.REPORT_TYPES)

    _print_banner("ÖZET")
    print(f"Örneklem: {len(companies)} şirket × {len(config.SCAN_YEARS)} yıl × {len(config.REPORT_TYPES)} rapor türü = {beklenen_toplam} kombinasyon")
    print(f"İndirildi : {indirildi}")
    print(f"Rapor Yok : {rapor_yok}")
    print(f"Hata      : {hata}")
    if indirilecek:
        # Beklenmiyor: her kombinasyon run() sonunda nihai bir duruma
        # (İndirildi/Rapor Yok/Hata) ulaşmalı. Görülüyorsa downloader.run()
        # bir kombinasyonu hiç işlemeden "İndirilecek" yer tutucusunda
        # bırakmış demektir (bkz. src/downloader.py'deki stuck_restored).
        print(f"UYARI     : {indirilecek} kombinasyon hâlâ 'İndirilecek' durumunda kaldı (beklenmiyordu)")
    print(f"Çıktı     : {config.OUTPUT_XLSX}")


if __name__ == "__main__":
    main()
