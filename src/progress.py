"""Parça (chunk) bazlı çalıştırma ve kaldığı yerden devam etme.

Tüm pipeline adımları (inventory, downloader, text_miner+classifier) bu
modül üzerinden 15-20 şirketlik gruplar halinde çalışır. Her parçadan sonra
durum diske yazılır; bir sonraki çalıştırma zaten tamamlanmış şirketleri
atlar. Bir şirkette hata oluşursa o şirket "hata" olarak işaretlenir ve akış
durmadan devam eder.
"""
from __future__ import annotations

import json
import logging
from collections.abc import Callable, Iterable, Sequence
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)


def load_json(path: Path, default: Any) -> Any:
    if not path.exists():
        return default
    try:
        with path.open("r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError) as exc:
        logger.warning("Bozuk durum dosyası %s okunamadı (%s), sıfırdan başlanıyor.", path, exc)
        return default


def save_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = path.with_suffix(path.suffix + ".tmp")
    with tmp_path.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, sort_keys=True)
    tmp_path.replace(path)  # atomik yazma: yarım kalmış dosya riskini önler


def chunked(items: Sequence[Any], size: int) -> Iterable[Sequence[Any]]:
    for i in range(0, len(items), size):
        yield items[i : i + size]


@dataclass
class StageProgress:
    """Tek bir pipeline adımı (örn. "inventory", "download") için durum.

    completed: başarıyla işlenmiş anahtarlar (örn. şirket kodu ya da
        "KOD|YIL|RAPOR_TURU" gibi birleşik anahtarlar).
    errors: anahtar -> son hata mesajı.
    """

    stage: str
    completed: set[str] = field(default_factory=set)
    errors: dict[str, str] = field(default_factory=dict)
    last_updated: str | None = None

    def is_done(self, key: str) -> bool:
        return key in self.completed

    def mark_done(self, key: str) -> None:
        self.completed.add(key)
        self.errors.pop(key, None)

    def mark_error(self, key: str, message: str) -> None:
        self.errors[key] = message

    def to_dict(self) -> dict:
        return {
            "stage": self.stage,
            "completed": sorted(self.completed),
            "errors": self.errors,
            "last_updated": self.last_updated,
        }

    @classmethod
    def from_dict(cls, stage: str, data: dict | None) -> "StageProgress":
        data = data or {}
        return cls(
            stage=stage,
            completed=set(data.get("completed", [])),
            errors=dict(data.get("errors", {})),
            last_updated=data.get("last_updated"),
        )


class ProgressStore:
    """progress.json içindeki tüm adımların durumunu tutar."""

    def __init__(self, path: Path):
        self.path = path
        self._raw: dict = load_json(path, {})

    def stage(self, name: str) -> StageProgress:
        return StageProgress.from_dict(name, self._raw.get(name))

    def save_stage(self, progress: StageProgress) -> None:
        progress.last_updated = datetime.now(timezone.utc).isoformat()
        self._raw[progress.stage] = progress.to_dict()
        save_json(self.path, self._raw)


def run_in_chunks(
    items: Sequence[Any],
    key_fn: Callable[[Any], str],
    process_fn: Callable[[Any], None],
    progress_store: ProgressStore,
    stage_name: str,
    chunk_size: int,
    on_chunk_done: Callable[[int, int, StageProgress], None] | None = None,
) -> StageProgress:
    """items'ı chunk_size'lık parçalar halinde işler.

    - Zaten tamamlanmış (progress'te "completed" olan) öğeler atlanır.
    - process_fn bir öğede istisna (exception) fırlatırsa, TÜM İŞLEM
      DURDURULMAZ: öğe "hata" olarak işaretlenir ve akış devam eder.
    - Her parça sonunda ilerleme diske kaydedilir (kaldığın yerden devam
      edebilmek için) ve on_chunk_done çağrılırsa özet bilgi iletilir.
    """
    progress = progress_store.stage(stage_name)
    total_chunks = (len(items) + chunk_size - 1) // chunk_size or 1

    for chunk_idx, chunk in enumerate(chunked(items, chunk_size), start=1):
        for item in chunk:
            key = key_fn(item)
            if progress.is_done(key):
                continue
            try:
                process_fn(item)
            except Exception as exc:  # noqa: BLE001 - kasıtlı: bir hata tüm taramayı durdurmasın
                logger.exception("'%s' işlenirken hata (%s aşaması)", key, stage_name)
                progress.mark_error(key, f"{type(exc).__name__}: {exc}")
            else:
                progress.mark_done(key)

        progress_store.save_stage(progress)
        if on_chunk_done is not None:
            on_chunk_done(chunk_idx, total_chunks, progress)

    return progress
