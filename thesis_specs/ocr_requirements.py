# -*- coding: utf-8 -*-
"""
OCR 扫描版《毕业论文（设计）指导规范》.pdf
使用 RapidOCR（ONNX），对中文印刷体识别较稳定；结果写入 UTF-8 文本。
"""
from __future__ import annotations

import os
import sys

import cv2
import fitz
import numpy as np
from rapidocr_onnxruntime import RapidOCR

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PDF = os.path.join(ROOT, "thesis_requirements", "毕业论文（设计）指导规范.pdf")
OUT = os.path.join(ROOT, "thesis_specs", "指导规范_OCR全文.txt")


def page_to_bgr(page: fitz.Page, zoom: float = 3.0) -> np.ndarray:
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    h, w = pix.height, pix.width
    rgb = np.frombuffer(pix.samples, dtype=np.uint8).reshape(h, w, 3)
    return cv2.cvtColor(rgb, cv2.COLOR_RGB2BGR)


def main() -> None:
    if not os.path.isfile(PDF):
        sys.stderr.write(f"missing: {PDF}\n")
        sys.exit(1)

    ocr = RapidOCR()
    doc = fitz.open(PDF)
    chunks: list[str] = []
    try:
        for i in range(len(doc)):
            img = page_to_bgr(doc[i], zoom=3.0)
            # 直接传 ndarray，避免含中文路径时落盘失败
            res, _elapsed = ocr(img)
            lines = [str(x[1]).strip() for x in (res or []) if x and len(x) >= 2 and x[1]]
            text = "\n".join(lines)
            chunks.append(f"\n\n===== 第 {i + 1} 页 / 共 {len(doc)} 页 =====\n\n")
            chunks.append(text + "\n")
            print(f"page {i + 1}/{len(doc)} lines {len(lines)}", flush=True)
    finally:
        doc.close()

    body = "".join(chunks).strip() + "\n"
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        f.write(body)
    print(OUT, "chars", len(body))


if __name__ == "__main__":
    main()
