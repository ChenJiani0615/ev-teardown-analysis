#!/usr/bin/env python3
"""part_table_check.py 回归测试（仅标准库，无第三方依赖）。

用法：
    python tests/test_part_table_check.py            # 校验，不一致则退出码 1
    python tests/test_part_table_check.py --update   # 按当前脚本行为重写基线

约定：
    tests/expected/<name>.json / <name>.md 的输入来自 samples/<name>.csv
    或 tests/fixtures/<name>.csv。基线只在确认"行为变更是预期的"之后更新，
    否则应视为回归。

比较口径：
    * JSON 走解析后深度比较（不受空行、缩进、行尾差异影响）；
    * Markdown 走逐行比较（仅归一化 CRLF），因为 md 里含手写文案，
      文案变化属于需要人看过的变更。
"""

from __future__ import annotations

import argparse
import difflib
import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / ".github" / "skills" / "ev-teardown-benchmark" / "scripts" / "part_table_check.py"
EXPECTED_DIR = ROOT / "tests" / "expected"
INPUT_DIRS = (ROOT / "samples", ROOT / "tests" / "fixtures")


def find_input(name: str) -> Path:
    for folder in INPUT_DIRS:
        for suffix in (".csv", ".xlsx", ".xlsm"):
            candidate = folder / f"{name}{suffix}"
            if candidate.exists():
                return candidate
    raise SystemExit(f"找不到 {name} 的输入清单（查找目录：{', '.join(str(d) for d in INPUT_DIRS)}）")


def run(script: Path, source: Path, fmt: str) -> str:
    proc = subprocess.run(
        [sys.executable, str(script), str(source), "--format", fmt],
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    if proc.returncode != 0:
        raise SystemExit(
            f"脚本执行失败（{source.name}, --format {fmt}）\n退出码 {proc.returncode}\n{proc.stderr}"
        )
    return proc.stdout


def normalize_md(text: str) -> str:
    """统一行尾与结尾换行：stdout 经 print() 会多一个换行，--out 写文件不会。"""
    return text.replace("\r\n", "\n").rstrip("\n") + "\n"


def diff_lines(expected: str, actual: str, limit: int = 40) -> str:

    lines = list(
        difflib.unified_diff(
            expected.splitlines(),
            actual.splitlines(),
            fromfile="expected",
            tofile="actual",
            lineterm="",
            n=2,
        )
    )
    if len(lines) > limit:
        lines = lines[:limit] + [f"... 还有 {len(lines) - limit} 行差异未显示"]
    return "\n".join(lines)


def check(update: bool) -> int:
    if not SCRIPT.exists():
        raise SystemExit(f"找不到被测脚本：{SCRIPT}")

    baselines = sorted(EXPECTED_DIR.glob("*.json"))
    if not baselines:
        raise SystemExit(f"{EXPECTED_DIR} 下没有基线文件，先运行 --update 生成")

    failures: list[str] = []
    for json_path in baselines:
        name = json_path.stem
        source = find_input(name)
        md_path = EXPECTED_DIR / f"{name}.md"

        actual_json = json.loads(run(SCRIPT, source, "json"))
        actual_md = normalize_md(run(SCRIPT, source, "md"))

        if update:
            json_path.write_text(
                json.dumps(actual_json, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
            )
            md_path.write_text(actual_md, encoding="utf-8")
            print(f"[updated] {name}  <- {source.relative_to(ROOT)}")
            continue

        expected_md = normalize_md(md_path.read_text(encoding="utf-8")) if md_path.exists() else ""
        status: list[str] = []

        if not md_path.exists():
            status.append("缺少 md 基线")
            failures.append(f"{name}: 缺少 {md_path.name}")

        if actual_json != json.loads(json_path.read_text(encoding="utf-8")):
            status.append("JSON 不一致")
            failures.append(f"{name}: JSON 基线不匹配")
            expected_json = json.loads(json_path.read_text(encoding="utf-8"))
            failures.append(diff_lines(
                json.dumps(expected_json, ensure_ascii=False, indent=2),
                json.dumps(actual_json, ensure_ascii=False, indent=2),
            ))

        if md_path.exists() and actual_md != expected_md:
            status.append("MD 不一致")
            failures.append(f"{name}: Markdown 基线不匹配")
            failures.append(diff_lines(expected_md, actual_md))

        mark = "ok" if not status else "FAIL"
        print(f"[{mark}] {name}  <- {source.relative_to(ROOT)}  {''.join(status)}")

    if update:
        print("基线已更新。请在提交前人工确认 diff 是预期变更。")
        return 0

    if failures:
        print("\n--- 回归失败 ---")
        for block in failures:
            print(block)
        return 1

    print(f"\n全部通过（{len(baselines)} 份基线）。")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="part_table_check.py 回归测试")
    parser.add_argument(
        "--update", action="store_true", help="按当前脚本输出重写 tests/expected 下的基线"
    )
    args = parser.parse_args()
    return check(args.update)


if __name__ == "__main__":
    raise SystemExit(main())
