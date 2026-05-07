"""Build author-voice-extended.md from skill参考论文/extracted_*.txt only."""
from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[4]
SKILL = ROOT / ".cursor/skills/chinese-thesis-aigc-mitigation"
SKILL_PAPER = ROOT / "skill参考论文"
OUT = SKILL / "author-voice-extended.md"

EXTRACTED_ORDER = [
    "extracted_期末论文.txt",
    "extracted_长文终稿.txt",
    "extracted_乡村振兴面临的主要问题_孟帅-2021010910_.txt",
    "extracted_意大利台地园及其代表人物作品简析_.txt",
    "extracted_pre讲稿.txt",
]


def main() -> None:
    parts: list[str] = [
        "# 笔调参考：五篇原文全文（仅 `skill参考论文/`）\n",
        "本文档与 `skill参考论文/extracted_*.txt` **同源**（由 `extract_skill_docx.py` 从对应 docx 生成）。"
        "仅供 Agent 对齐用户既有写作节律与信息密度；**勿整段照抄**到待检测文本。\n",
        "\n---\n",
    ]
    for name in EXTRACTED_ORDER:
        p = SKILL_PAPER / name
        if not p.is_file():
            continue
        parts.append(f"\n## 全文：`skill参考论文/{name}`\n\n")
        parts.append(p.read_text(encoding="utf-8").rstrip())
        parts.append("\n\n---\n")

    OUT.write_text("".join(parts), encoding="utf-8")
    print("Wrote", OUT.relative_to(ROOT), "chars", OUT.stat().st_size)


if __name__ == "__main__":
    main()
