# -*- coding: utf-8 -*-
"""
bjfu-meetingppt 质检脚本
用法：
  python qa_deck.py check  <pptx路径>              # 越界/重叠检查（python-pptx）
  python qa_deck.py pixels <渲染PNG目录>           # 逐页内容覆盖率/空白格子分析
检查通过 ≠ 视觉合格；几何通过后仍需渲染并做视觉评审。
"""
import sys

from pptx import Presentation
from pptx.util import Emu

SW, SH = 13.333, 7.5


def check(pptx_path):
    p = Presentation(pptx_path)
    issues = []
    for si, s in enumerate(p.slides, 1):
        pics, txts = [], []
        for sh in s.shapes:
            if sh.left is None:
                continue
            l, t = Emu(sh.left).inches, Emu(sh.top).inches
            r, b = l + Emu(sh.width).inches, t + Emu(sh.height).inches
            if l < -0.02 or t < -0.02 or r > SW + 0.02 or b > SH + 0.02:
                issues.append(f"slide{si} '{sh.name}' 越界 L{l:.2f} T{t:.2f} R{r:.2f} B{b:.2f}")
            box = (l, t, r, b)
            if sh.shape_type == 13:
                pics.append((sh.name, box))
            elif sh.has_text_frame and sh.text_frame.text.strip():
                txts.append((sh.text_frame.text[:12], box))
        for i in range(len(pics)):
            for j in range(i + 1, len(pics)):
                a, b2 = pics[i][1], pics[j][1]
                ox = max(0, min(a[2], b2[2]) - max(a[0], b2[0]))
                oy = max(0, min(a[3], b2[3]) - max(a[1], b2[1]))
                if ox > 0.05 and oy > 0.05:
                    issues.append(f"slide{si} 图片重叠 {pics[i][0]}×{pics[j][0]} ({ox:.2f}×{oy:.2f}in)")
        for tname, tb in txts:
            for pname, pb in pics:
                ox = max(0, min(tb[2], pb[2]) - max(tb[0], pb[0]))
                oy = max(0, min(tb[3], pb[3]) - max(tb[1], pb[1]))
                if ox > 0.05 and oy > 0.05:
                    issues.append(f"slide{si} 文字[{tname}]与图片重叠 ({ox:.2f}×{oy:.2f}in)")
    print(f"slides={len(p.slides._sldIdLst)}  bounds/overlap issues={len(issues)}")
    for i in issues:
        print(" -", i)
    return issues


def pixels(png_dir):
    import os
    import numpy as np
    from PIL import Image
    warn = []
    for f in sorted(os.listdir(png_dir), key=lambda n: (len(n), n)):
        if not f.lower().endswith(".png"):
            continue
        im = Image.open(os.path.join(png_dir, f)).convert("RGB")
        a = np.asarray(im, dtype=np.int16)
        h, w, _ = a.shape
        nonwhite = (np.abs(a - 255).max(axis=2) > 12)
        cover = nonwhite.mean()
        empties = []
        for gy in range(4):
            for gx in range(6):
                blk = nonwhite[gy * h // 4:(gy + 1) * h // 4, gx * w // 6:(gx + 1) * w // 6]
                if blk.mean() < 0.02:
                    empties.append((gy, gx))
        flag = " <-- 检查是否大面积空白" if len(empties) >= 5 else ""
        print(f"{f:12s} 覆盖率{cover * 100:5.1f}%  空白格{len(empties)}/24{flag}")
        if len(empties) >= 5:
            warn.append(f)
    return warn


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else ""
    if cmd == "check":
        check(sys.argv[2])
    elif cmd == "pixels":
        pixels(sys.argv[2])
    else:
        print(__doc__)
