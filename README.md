# qingtong_paper

Undergraduate **environmental design** thesis materials: interior design for a tea–coffee venue in Pengzhou (Longxing Temple waterfront district). Repository text is mostly **Chinese**; **top-level folders use English names** for tooling and portability.

## Repository layout

| Folder | Purpose |
|--------|---------|
| `final_thesis/` | Final thesis Markdown, body source, and `build_final_thesis.py` (exports Word/PDF; reads images from `drawings/`). |
| `proposal/` | Opening report (开题) drafts and notes. |
| `thesis_specs/` | School writing rules: OCR excerpt (`规范依据_OCR摘录.md`), full OCR text, and `ocr_requirements.py` to re-run OCR. |
| `thesis_requirements/` | Source PDF for OCR (e.g. 《毕业论文（设计）指导规范》). |
| `drawings/` | Project drawings referenced by the thesis Markdown (`../drawings/...`). |
| `detection_reports/` | AIGC / similarity-style PDF reports for archival reference. |
| `skill_reference_papers/` | Reference `.docx` files and `extracted_*.txt` used by the Cursor skill `chinese-thesis-aigc-mitigation`. |
| `参考设计/` | Optional reference PDFs from other projects (rename to `reference_designs` locally if you prefer English-only paths). |
| `.cursor/skills/chinese-thesis-aigc-mitigation/` | Cursor Agent Skill: mitigating AIGC-style detection in Chinese academic writing. |

## Scripts (common)

From the **repository root**:

```bash
# Re-extract plain text from docx in skill_reference_papers/
python .cursor/skills/chinese-thesis-aigc-mitigation/scripts/extract_skill_docx.py
python .cursor/skills/chinese-thesis-aigc-mitigation/scripts/build_author_voice_extended.py
```

```bash
# OCR the school guideline PDF into thesis_specs/ (requires rapidocr, pymupdf, opencv, numpy)
python thesis_specs/ocr_requirements.py
```

```bash
# Build Word/PDF from final_thesis Markdown (requires python-docx and a Word-to-PDF path if applicable)
python final_thesis/build_final_thesis.py
```

## Notes

- File names inside folders (e.g. Chinese thesis titles) are unchanged; only **directory names** were normalized to English.
- `_paths.txt` is a flat index of asset paths; `_full_tree.txt` may be an older tree snapshot.
- Always follow your school’s **latest** printed guidelines; OCR and excerpts here are aids, not a substitute for official documents.
