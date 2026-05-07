"""Extract plain text from docx in skill参考论文 for voice samples."""
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def docx_plain(path: Path) -> str:
    with zipfile.ZipFile(path, "r") as z:
        root = ET.fromstring(z.read("word/document.xml"))
    parts: list[str] = []
    for p in root.iter(f"{W}p"):
        texts = [t.text or "" for t in p.iter(f"{W}t")]
        line = "".join(texts).strip()
        if line:
            parts.append(line)
    return "\n".join(parts)


def main() -> None:
    # .../qingtong_paper/.cursor/skills/.../scripts/this.py -> repo root = parents[4]
    root = Path(__file__).resolve().parents[4]
    folder = root / "skill参考论文"
    if not folder.is_dir():
        raise SystemExit(f"Missing folder: {folder}")
    # 纯文本与 docx 同目录，便于对照与手工裁剪进 author-voice-sample.md
    out_dir = folder
    for f in sorted(folder.glob("*.docx")):
        text = docx_plain(f)
        safe = "".join(c if c.isalnum() or c in "._-" else "_" for c in f.stem)[:80]
        out = out_dir / f"extracted_{safe}.txt"
        out.write_text(text, encoding="utf-8")
        rel = out.relative_to(root)
        print(f.name, "->", rel.as_posix(), "chars", len(text))


if __name__ == "__main__":
    main()
