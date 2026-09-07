"""模板配置加载。

字段映射全部放在 tradecheck/config/template_*.json，代码里不写死任何表头文字。
"""

from __future__ import annotations

import json
import os
from typing import Any, Dict, List

CONFIG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config")


class TemplateError(Exception):
    pass


def list_templates() -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    if not os.path.isdir(CONFIG_DIR):
        return out
    for fn in sorted(os.listdir(CONFIG_DIR)):
        if fn.startswith("template_") and fn.endswith(".json"):
            try:
                cfg = load_template_file(os.path.join(CONFIG_DIR, fn))
            except TemplateError:
                continue
            out.append(
                {
                    "template_id": cfg["template_id"],
                    "display_name": cfg.get("display_name", cfg["template_id"]),
                    "description": cfg.get("description", ""),
                }
            )
    return out


def load_template_file(path: str) -> Dict[str, Any]:
    try:
        with open(path, "r", encoding="utf-8") as fh:
            cfg = json.load(fh)
    except FileNotFoundError:
        raise TemplateError(f"模板配置文件不存在：{path}")
    except json.JSONDecodeError as exc:
        raise TemplateError(f"模板配置不是合法 JSON：{os.path.basename(path)}（{exc}）")
    if not isinstance(cfg, dict) or "template_id" not in cfg:
        raise TemplateError(f"模板配置缺少 template_id：{os.path.basename(path)}")
    for role in ("invoice", "packing"):
        if role not in cfg.get("documents", {}):
            raise TemplateError(f"模板 {cfg['template_id']} 缺少 documents.{role} 配置")
    return cfg


def load_template(template_id: str) -> Dict[str, Any]:
    tid = str(template_id).strip().upper()
    path = os.path.join(CONFIG_DIR, f"template_{tid.casefold()}.json")
    if not os.path.isfile(path):
        available = "、".join(t["template_id"] for t in list_templates()) or "（无）"
        raise TemplateError(f"未找到模板 {tid}。可用模板：{available}")
    cfg = load_template_file(path)
    if str(cfg["template_id"]).upper() != tid:
        raise TemplateError(
            f"模板文件 {os.path.basename(path)} 内部 template_id={cfg['template_id']}，与请求的 {tid} 不一致"
        )
    return cfg


def money_config(template: Dict[str, Any]) -> Dict[str, Any]:
    cfg = template.get("money") or {}
    rounding = cfg.get("rounding") or {}
    tolerance = cfg.get("tolerance") or {}
    return {
        "default_currency": cfg.get("default_currency"),
        "digits": int(rounding.get("digits", 2)),
        "rounding_mode": rounding.get("mode", "half_up"),
        "tolerance_absolute": str(tolerance.get("absolute", "0.01")),
        "note": cfg.get("note", ""),
    }


# 别名：界面层使用 available_templates，逻辑层保留 list_templates 语义。
available_templates = list_templates
