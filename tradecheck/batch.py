"""JSON-manifest batch checking for the TradeCheck command line."""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

from .models import STATUS_DIFF, STATUS_SKIP, STATUS_TODO
from .pipeline import PipelineError, check_pair, extract_pair
from .templates import TemplateError, load_template
from .version import TOOL_NAME, VERSION


class BatchManifestError(ValueError):
    """The manifest cannot be used to start a batch run."""


def _load_manifest(path: Path) -> Tuple[Dict[str, Any], Set[Path]]:
    try:
        raw = path.read_text(encoding="utf-8")
    except OSError as exc:
        raise BatchManifestError(f"无法读取清单文件：{exc}") from exc
    try:
        manifest = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise BatchManifestError(
            f"清单不是合法 JSON（第 {exc.lineno} 行，第 {exc.colno} 列）：{exc.msg}"
        ) from exc

    if not isinstance(manifest, dict):
        raise BatchManifestError("清单最外层必须是 JSON 对象。")

    default_template = manifest.get("template", "A")
    if not isinstance(default_template, str) or not default_template.strip():
        raise BatchManifestError("清单字段 template 必须是非空字符串。")
    try:
        load_template(default_template)
    except TemplateError as exc:
        raise BatchManifestError(str(exc)) from exc

    pairs = manifest.get("pairs")
    if not isinstance(pairs, list) or not pairs:
        raise BatchManifestError("清单必须包含非空 pairs 数组。")

    protected_paths: Set[Path] = {path.resolve()}
    for pair in pairs:
        if not isinstance(pair, dict):
            continue
        for key in ("invoice", "packing"):
            value = pair.get(key)
            if isinstance(value, str) and value.strip():
                file_path = Path(value).expanduser()
                if not file_path.is_absolute():
                    file_path = path.parent / file_path
                protected_paths.add(file_path.resolve())
    return manifest, protected_paths


def _pair_path(value: Any, key: str, manifest_dir: Path) -> Path:
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"字段 {key} 必须是非空文件路径。")
    path = Path(value).expanduser()
    if not path.is_absolute():
        path = manifest_dir / path
    return path.resolve()


def run_batch(manifest_path: str) -> Tuple[Dict[str, Any], int, Set[Path]]:
    """Run each manifest pair independently.

    Returns the JSON-compatible report, its suggested process exit code, and
    paths that must never be selected as the output destination.
    """
    path = Path(manifest_path).expanduser().resolve()
    manifest, protected_paths = _load_manifest(path)
    manifest_dir = path.parent
    default_template = manifest.get("template", "A").strip().upper()
    results: List[Dict[str, Any]] = []
    seen_ids: Set[str] = set()
    passed = 0
    review_required = 0
    input_errors = 0

    for index, pair in enumerate(manifest["pairs"], start=1):
        default_id = f"pair-{index:03d}"
        pair_id: Any = pair.get("id", default_id) if isinstance(pair, dict) else default_id
        if not isinstance(pair_id, str) or not pair_id.strip():
            pair_id = default_id
            validation_error = "字段 id 必须是非空字符串。"
        else:
            pair_id = pair_id.strip()
            validation_error = ""
        if pair_id in seen_ids:
            validation_error = validation_error or f"批次 id 重复：{pair_id}。每组应使用唯一 id。"
        seen_ids.add(pair_id)

        requested_template = (
            pair.get("template", default_template) if isinstance(pair, dict) else default_template
        )
        if not isinstance(requested_template, str) or not requested_template.strip():
            requested_template = default_template
            validation_error = validation_error or "字段 template 必须是非空字符串。"
        requested_template = requested_template.strip().upper()

        invoice_path: Optional[Path] = None
        packing_path: Optional[Path] = None
        try:
            if not isinstance(pair, dict):
                raise ValueError("每个 pairs 项都必须是 JSON 对象。")
            if validation_error:
                raise ValueError(validation_error)
            invoice_path = _pair_path(pair.get("invoice"), "invoice", manifest_dir)
            packing_path = _pair_path(pair.get("packing"), "packing", manifest_dir)
            if invoice_path == packing_path:
                raise ValueError("同一组的 invoice 和 packing 必须指向不同文件。")
            invoice, packing, template = extract_pair(
                str(invoice_path),
                str(packing_path),
                requested_template,
                invoice_name=invoice_path.name,
                packing_name=packing_path.name,
            )
            result, extra = check_pair(invoice, packing, template)
            overall = result.overall
            needs_review = any(
                result.count(status) for status in (STATUS_DIFF, STATUS_TODO, STATUS_SKIP)
            ) or bool(result.unchecked)
            if not needs_review:
                passed += 1
            else:
                review_required += 1
            results.append(
                {
                    "id": pair_id,
                    "status": "checked",
                    "template": requested_template,
                    "invoice": {
                        "file_name": invoice.file_name,
                        "sha256": invoice.sha256,
                        "sheet": invoice.sheet,
                    },
                    "packing": {
                        "file_name": packing.file_name,
                        "sha256": packing.sha256,
                        "sheet": packing.sheet,
                    },
                    "overall": overall,
                    "overall_note": result.overall_note,
                    "summary": result.summary(),
                    "coverage": list(result.coverage),
                    "unchecked": list(result.unchecked),
                    "currency": extra,
                    "findings": [finding.to_dict() for finding in result.findings],
                }
            )
        except (PipelineError, TemplateError, OSError, ValueError) as exc:
            input_errors += 1
            results.append(
                {
                    "id": pair_id,
                    "status": "error",
                    "template": requested_template,
                    "invoice_file": invoice_path.name if invoice_path else _safe_basename(pair, "invoice"),
                    "packing_file": packing_path.name if packing_path else _safe_basename(pair, "packing"),
                    "error_type": type(exc).__name__,
                    "error": str(exc),
                }
            )
        except Exception as exc:  # Keep later pairs running if one input exposes an unexpected parser error.
            input_errors += 1
            results.append(
                {
                    "id": pair_id,
                    "status": "error",
                    "template": requested_template,
                    "invoice_file": invoice_path.name if invoice_path else _safe_basename(pair, "invoice"),
                    "packing_file": packing_path.name if packing_path else _safe_basename(pair, "packing"),
                    "error_type": type(exc).__name__,
                    "error": "处理该组单据时发生未预期错误，请单独重试并联系维护者。",
                }
            )

    if input_errors:
        batch_status = "completed_with_errors"
    elif review_required:
        batch_status = "review_required"
    else:
        batch_status = "completed"

    report = {
        "schema": "tradecheck.batch.v1",
        "tool": {"name": TOOL_NAME, "version": VERSION},
        "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "status": batch_status,
        "counts": {
            "total": len(results),
            "passed": passed,
            "review_required": review_required,
            "input_errors": input_errors,
        },
        "results": results,
    }
    exit_code = 0 if batch_status == "completed" else 1
    return report, exit_code, protected_paths


def _safe_basename(pair: Any, key: str) -> str:
    if isinstance(pair, dict) and isinstance(pair.get(key), str):
        return Path(pair[key]).name
    return ""


def write_report(report: Dict[str, Any], output: str, force: bool, protected_paths: Set[Path]) -> None:
    payload = json.dumps(report, ensure_ascii=False, indent=2) + "\n"
    if output == "-":
        print(payload, end="")
        return

    output_path = Path(output).expanduser().resolve()
    if output_path in protected_paths:
        raise OSError("JSON 输出路径与清单或输入工作簿相同，已拒绝写入。")
    if output_path.exists() and not force:
        raise FileExistsError(f"输出文件已存在：{output_path.name}。如需覆盖，请明确使用 --force。")

    mode = "w" if force else "x"
    with output_path.open(mode, encoding="utf-8", newline="\n") as handle:
        handle.write(payload)
