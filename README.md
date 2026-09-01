# BIST 100 Kapsam 3 (Scope 3) Emisyon Raporlama Olgunluğu Tarama Ajanı

Akademik bir makale için, **SABİT bir BIST 100 örnekleminin** 2023-2025 mali
yıllarına ait sürdürülebilirlik raporları ve faaliyet raporlarını tarayıp
Kapsam 3 emisyon açıklamalarını ve olgunluk düzeyini tespit eden, sayfa
numarası ve doğrudan alıntı ile denetlenebilir bir Excel veri seti üreten
modüler bir Python ajanı.

## ⚠️ Ağ erişimi hakkında önemli not

Bu iskelet, üzerinde geliştirildiği kapalı ağ ortamında (Claude Code
sandbox) **kgk.gov.tr ve şirket web sitelerine erişemiyor** (kurumsal
proxy politikası HTTP 403 ile engelliyor). Bu nedenle:

- `inventory.py`'nin otomatik site keşfi ve `downloader.py`'nin PDF indirme
  adımları **bu ortamda test EDİLEMEMİŞTİR** — mantık yazılmış ve
  hatalara dayanıklı (retry, timeout, tek-şirket-hatası-akışı-durdurmaz)
  şekilde tasarlanmıştır, ama gerçek bir ağ isteği bu kutuda hiç
  yapılamamıştır.
- Ağ GEREKTİRMEYEN her şey (xlsx okuma, şirket/kod eşleştirme, PDF metin
  madenciliği, olgunluk sınıflandırması, Excel çıktısı) **sentetik
  PDF'lerle uçtan uca test edilmiş ve doğrulanmıştır** (`tests/test_pipeline_smoke.py`).
- Ajanı gerçek verilerle çalıştırmak için **internet erişimi olan bir
  ortamda** (kendi bilgisayarınız, farklı bir Claude Code oturumu vb.)
  çalıştırmanız gerekir.

## Mimari

```
config.py                  Merkezi ayarlar: yollar, anahtar kelimeler, chunk boyutu
main.py                    Orkestratör — tüm pipeline'ı sırayla/parça parça çalıştırır
src/
  models.py                Company, ReportRef, Finding, MaturityResult veri modelleri
  progress.py               Parça (chunk) bazlı ilerleme takibi + kaldığı yerden devam
  inventory.py              Şirket listesi + IR/resmi site URL keşfi (cache'li)
  downloader.py             KGK + şirket IR sitesinden PDF indirme (idempotent)
  text_miner.py             PyMuPDF ile sayfa sayfa Kapsam 3 terim taraması
  maturity_classifier.py    Kural tabanlı olgunluk etiketleme (şeffaf, denetlenebilir)
  excel_writer.py           4 sekmeli denetlenebilir Excel çıktısı
tools/
  compare_tsrs_bist100.py   Tek seferlik: mevcut 150 raporluk KGK envanterini
                             BIST100 örneklemiyle karşılaştırır (ağ gerektirmez)
tests/
  test_pipeline_smoke.py    Ağsız uçtan uca doğrulama (sentetik PDF ile)
data/
  input/                    Girdi referans dosyaları (BIST100 listesi, 788 şirket
                             referansı, TSRS/KGK rapor envanteri — kullanıcı sağladı)
  cache/                    inventory.py / downloader.py / text_miner.py önbellekleri
  downloads/                İndirilen PDF'ler (kod/yıl/rapor_türü.pdf)
  output/                   Nihai Excel çıktısı
  progress.json             Kaldığın yerden devam etme durumu
```

### Veri akışı

```
BIST100 listesi (100, sabit) ──┐
788 şirket referansı ──────────┼─▶ inventory.py ──▶ company_websites.json (cache)
                                │
KGK/TSRS envanteri (150 rapor) ┴─▶ downloader.py ──▶ data/downloads/KOD/YIL/tür.pdf
                                     (+ IR sitesi taraması)   + report_manifest.json
                                              │
                                              ▼
                                       text_miner.py ──▶ findings.json
                                     (sayfa no + alıntı)
                                              │
                                              ▼
                                  maturity_classifier.py
                                   (kural tabanlı etiketleme)
                                              │
                                              ▼
                                      excel_writer.py
                              ──▶ data/output/kapsam3_bulgular.xlsx
                                   (4 sekme, hepsi denetlenebilir)
```

## Kurulum

```bash
pip install -r requirements.txt
```

## Çalıştırma

```bash
# Tüm pipeline (inventory -> download -> mine -> report), 15 şirketlik parçalar halinde
python main.py

# Sadece belirli bir aşama
python main.py --stage inventory
python main.py --stage download
python main.py --stage mine
python main.py --stage report

# Parça boyutunu değiştir
python main.py --chunk-size 20
```

Her aşama `data/progress.json`'a ilerlemesini kaydeder. Betik tekrar
çalıştırıldığında zaten tamamlanmış şirketleri/raporları ATLAR — kaldığı
yerden devam eder. Bir şirkette/raporda hata olursa **tüm işlem durmaz**;
o kayıt "Hata" olarak işaretlenir ve akış devam eder. Her parça sonunda
konsola kısa bir özet basılır.

### Tek seferlik: mevcut envanteri BIST100 ile karşılaştır

```bash
python tools/compare_tsrs_bist100.py
```

Elinizdeki 150 satırlık KGK rapor envanterini SABİT BIST100 listesiyle
karşılaştırır: hangi envanterin BIST100 dışı olduğunu ve BIST100'de olup
envanterde hiç raporu bulunmayan şirketleri listeler (envanteri silmez,
sadece raporlar). Ağ erişimi gerektirmez.

### Testler

```bash
python tests/test_pipeline_smoke.py
```

## Çıktı: `data/output/kapsam3_bulgular.xlsx`

| Sekme | İçerik |
| --- | --- |
| **Ham Bulgular** | Her terim eşleşmesi: şirket, yıl, rapor türü, sayfa no, eşleşen terim, ~300 karakterlik bağlam alıntısı, olgunluk sinyali, kategori no |
| **Şirket-Yıl-RaporTürü Özeti** | 100 şirket × 3 yıl × 2 rapor türü TAM matris (600 satır) — rapor durumu, olgunluk etiketi, gerekçe, kaynak URL |
| **Kategori 1-15 Matrisi** | Hangi GHG Protokolü Kapsam 3 kategorisinin (1-15) hangi rapor için bulunduğu |
| **Eksik Raporlar** | Rapor bulunamayan/indirilemeyen her şirket-yıl-rapor türü kombinasyonu |

**Önemli:** Tüm olgunluk etiketleri kural tabanlı bir sezgisel yöntemle
(heuristic) otomatik üretilir ve her satırda AÇIKÇA **"ön etiket, manuel
teyit gerekir"** notuyla işaretlenir. Akademik yayında nihai veri olarak
kullanılmadan önce elle doğrulanmalıdır.

### Olgunluk etiketleri (öncelik sırasıyla, kural tabanlı)

1. **Metodolojili-Güvenceli** — "güvence" / "assurance" / "ISO 14064" geçiyorsa
2. **Kategori Bazlı** — "Kategori N" (N=1-15) geçiyorsa
3. **Toplam Sayısal** — "tCO2e" gibi bir emisyon birimi geçiyorsa
4. **Nitel** — yalnızca "Kapsam 3" / "Scope 3" / "Değer Zinciri" geçiyorsa
5. **Açıklama Yok** — rapor tarandı ama hiçbir sinyal bulunamadı

## Girdi dosyaları (`data/input/`)

- `bist100_liste.xlsx` — SABİT 100 şirketlik örneklem (BIST_Kod, Sirket_Unvani, Sehir)
- `sirketler_referans.xlsx` — BIST'te işlem gören ~788 şirketin tam listesi
  (Kod, Ünvan, Şehir, Bağımsız Denetim Kuruluşu) — yalnızca kod/ünvan
  eşleştirme REFERANSI, örneklem bundan etkilenmez. **Not:** bazı satırlarda
  tek şirket için birden fazla kod virgülle ayrılmış tutulur (örn.
  `"GARAN, TGB"`, farklı pay sınıfları/eski kodlar) — `inventory.py` bunları
  ayrıştırıp her kodu ayrı ayrı indeksler.
- `tsrs_rapor_envanteri.xlsx` — önceden derlenmiş, KGK'da yayımlanan 150
  sürdürülebilirlik/entegre faaliyet raporu bağlantısı (2023-2025), BIST
  koduyla ön-eşleştirilmiş.
- `manual_ir_urls.json` (opsiyonel, siz oluşturursunuz) — elle doğrulanmış
  yatırımcı ilişkileri URL'leri. Format:
  ```json
  {
    "AKBNK": {
      "resmi_web_sitesi": "https://www.akbankinvestorrelations.com",
      "yatirimci_iliskileri_url": "https://www.akbankinvestorrelations.com/tr/",
      "dogrulandi": true
    }
  }
  ```
  `inventory.py` bu dosyadaki kayıtları otomatik keşiften önce kullanır.
