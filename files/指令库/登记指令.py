#!/usr/bin/env python3
"""Upsert a keyword command into 超级指令.json and regenerate *超级指令.docx."""

from __future__ import annotations

import argparse
import json
from datetime import date
from pathlib import Path

import importlib.util

ROOT = Path(__file__).resolve().parent
DATA_PATH = ROOT / "超级指令.json"


def _load_generate():
    script = ROOT / "更新超级指令.py"
    spec = importlib.util.spec_from_file_location("update_super_commands", script)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module.generate


def upsert(entry: dict) -> dict:
    data = json.loads(DATA_PATH.read_text(encoding="utf-8"))
    commands = data.setdefault("commands", [])
    keyword = " ".join(entry["keyword"].split())
    today = date.today().isoformat()
    existing = next(
        (item for item in commands if item["keyword"].casefold() == keyword.casefold()),
        None,
    )
    if existing:
        existing.update({key: value for key, value in entry.items() if value is not None})
        action = "更新"
    else:
        next_id = max((item.get("id", 0) for item in commands), default=0) + 1
        record = {
            "id": next_id,
            "keyword": keyword,
            "type": entry.get("type") or "关键词口令",
            "created": entry.get("created") or today,
            "source": entry.get("source") or "对话中新生成",
            "purpose": entry.get("purpose") or "",
            "inputs": entry.get("inputs") or "",
            "operation": entry.get("operation") or "",
            "outputs": entry.get("outputs") or "",
            "notes": entry.get("notes") or "",
        }
        commands.append(record)
        action = "新增"
    data["updated_at"] = today
    DATA_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    _load_generate()()
    print(f"{action}口令：{keyword}")
    return data


def main():
    parser = argparse.ArgumentParser(description="登记关键词新指令/提示词命令，并刷新 *超级指令.docx")
    parser.add_argument("--keyword", required=True)
    parser.add_argument("--type", default="关键词口令")
    parser.add_argument("--purpose", required=True)
    parser.add_argument("--inputs", default="")
    parser.add_argument("--operation", required=True)
    parser.add_argument("--outputs", default="")
    parser.add_argument("--source", default="对话中新生成")
    parser.add_argument("--notes", default="")
    args = parser.parse_args()
    upsert(vars(args))


if __name__ == "__main__":
    main()
