"""本地保存列映射模板和核对项目历史。"""

from __future__ import annotations

import hashlib
import json
import os
import uuid
from decimal import Decimal
from dataclasses import asdict, is_dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _file_record(path: str | Path) -> dict[str, Any]:
    resolved = Path(path).expanduser().resolve()
    stat = resolved.stat()
    return {
        "path": str(resolved),
        "size": stat.st_size,
        "modified_ns": stat.st_mtime_ns,
        "sha256": _sha256(resolved),
    }


def _jsonable(value: Any) -> Any:
    if is_dataclass(value):
        return _jsonable(asdict(value))
    if isinstance(value, dict):
        return {str(key): _jsonable(item) for key, item in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [_jsonable(item) for item in value]
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, Decimal):
        return str(value)
    if isinstance(value, (datetime,)):
        return value.isoformat()
    if hasattr(value, "value"):
        return value.value
    return value


class LocalProjectStore:
    """只追加项目记录；本轮不提供删除历史功能。"""

    def __init__(self, root: str | Path | None = None):
        base = Path(os.environ.get("LOCALAPPDATA", Path.home()))
        self.root = Path(root) if root is not None else base / "银行流水核对工具"
        self.mapping_path = self.root / "列映射模板.json"
        self.projects_dir = self.root / "项目历史"

    @staticmethod
    def mapping_fingerprint(source: str, columns: Iterable[Any]) -> str:
        normalized = sorted(str(column).strip() for column in columns)
        raw = json.dumps([source, normalized], ensure_ascii=False, separators=(",", ":"))
        return hashlib.sha256(raw.encode("utf-8")).hexdigest()

    def _read_mappings(self) -> dict[str, Any]:
        if not self.mapping_path.exists():
            return {}
        return json.loads(self.mapping_path.read_text(encoding="utf-8-sig"))

    def _write_json(self, path: Path, payload: Any) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        temporary = path.with_suffix(path.suffix + ".new")
        temporary.write_text(
            json.dumps(_jsonable(payload), ensure_ascii=False, indent=2),
            encoding="utf-8-sig",
        )
        os.replace(temporary, path)

    def save_mapping_template(self, source: str, columns: Iterable[Any], mapping: dict[str, Any], name: str) -> str:
        fingerprint = self.mapping_fingerprint(source, columns)
        records = self._read_mappings()
        records[fingerprint] = {
            "fingerprint": fingerprint,
            "source": source,
            "name": str(name),
            "columns": sorted(str(column).strip() for column in columns),
            "mapping": _jsonable(mapping),
            "updated_at": datetime.now().isoformat(timespec="seconds"),
        }
        self._write_json(self.mapping_path, records)
        return fingerprint

    def load_mapping_template(self, source: str, columns: Iterable[Any]) -> dict[str, Any] | None:
        return self._read_mappings().get(self.mapping_fingerprint(source, columns))

    def save_project(
        self,
        *,
        bank_path: str | Path,
        journal_path: str | Path,
        report_path: str | Path,
        bank_mapping: dict[str, Any],
        journal_mapping: dict[str, Any],
        parameters: Any,
        result_counts: dict[str, Any],
        version: str = "3.2",
    ) -> dict[str, Any]:
        created_at = datetime.now()
        project_id = created_at.strftime("%Y%m%d-%H%M%S-") + uuid.uuid4().hex[:8]
        record = {
            "project_id": project_id,
            "created_at": created_at.isoformat(timespec="seconds"),
            "version": version,
            "files": {
                "bank": _file_record(bank_path),
                "journal": _file_record(journal_path),
                "report": _file_record(report_path),
            },
            "mappings": {"bank": _jsonable(bank_mapping), "journal": _jsonable(journal_mapping)},
            "parameters": _jsonable(parameters),
            "result_counts": _jsonable(result_counts),
        }
        self._write_json(self.projects_dir / f"{project_id}.json", record)
        return record

    def list_projects(self) -> list[dict[str, Any]]:
        if not self.projects_dir.exists():
            return []
        records = []
        for path in self.projects_dir.glob("*.json"):
            try:
                records.append(json.loads(path.read_text(encoding="utf-8-sig")))
            except (OSError, ValueError):
                continue
        return sorted(records, key=lambda item: str(item.get("created_at", "")), reverse=True)

    def resolve_report(self, project_id: str, *, verify_hash: bool = True) -> Path:
        record_path = self.projects_dir / f"{project_id}.json"
        record = json.loads(record_path.read_text(encoding="utf-8-sig"))
        report = Path(record["files"]["report"]["path"])
        if not report.exists():
            raise FileNotFoundError(f"历史报告不存在：{report}")
        if verify_hash and _sha256(report) != record["files"]["report"]["sha256"]:
            raise ValueError("历史报告内容已经改变，原项目记录已失效")
        return report
