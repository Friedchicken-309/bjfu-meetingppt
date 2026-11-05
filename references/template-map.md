# 模板结构档案（assets/PPT模版.pptx）

母版无自有形状，装饰全部在各版式内；创建页面时装饰自动透出，无需重画。

## 画布与主题

- 16:9（13.333×7.5 in），主题"时尚混搭风"（蓝绿系）。
- 主题色：dk2=#5EA9C1，accent1=#3D6D84，accent2=#243848（藏青）。
- 中文字体：微软雅黑（模板示例页同款）；标题色为 `tx2 lumMod75%` ≈ **#467E90**。
- 编号块/强调色：**#31859C**；浅面板色推荐 #F1F6F8。

## 五个版式（索引即 Deck 内顺序）

| 索引 | 名称 | 关键占位符 | 版面要点 |
|---|---|---|---|
| 0 | 封面 | BODY ph idx=10（0.01, 2.12, 13.32×2.34） | 中部青蓝横带（y 2.12–4.46）；标题白字居中，副标题放 y≈3.80 白字居中 |
| 1 | 目录 | OBJECT ph idx=10（用于第 1 条） | 左侧深色图区（0–5.26in 勿放内容）；右侧 4 个编号块 #31859C（L=6.28, W=1.23, H=0.70，行 y≈1.29/2.52/3.86/5.09），编号右侧浅灰胶囊条 L≈7.79–12.37，**标题文字居中放在胶囊条上**（文字框 7.95, 4.25 宽，与行同高，anchor=ctr, algn=ctr） |
| 2 | 章节封面 | OBJECT ph idx=10（标题）、BODY ph idx=11（编号） | 右侧青色面板（x≥7.86）+ 超大数字（239pt, accent1, Arial，直接写入文本继承版式样式，勿自设字号）；左侧白区放标题，建议重定位到 (1.35, 3.90, 6.20×0.95) 28pt 藏青加粗 |
| 3 | 内容 | 仅页码 ph idx=12（12.24, 6.95） | 白底自由排版；内容区 y≈0.95–6.85、x 0.45–12.88；标题条=0.10×0.44in 青块 + 20pt 加粗 #467E90 |
| 4 | 空白 | 日期/页脚/页码 | 备用 |

目录行 y 值不要硬编码——从版式 shapes 中按 `Google Shape;397`（宽>1e6 EMU 且文本为数字）读取（ppt_kit.Deck.toc 已实现）。

## 换用其他模板时的勘察方法

运行以下勘察，替换本文档的对应数值，并同步修改 `ppt_kit.py` 的版式索引与占位符 idx：

```python
from pptx import Presentation
from pptx.util import Emu
p = Presentation("新模板.pptx")
print("size:", p.slide_width, p.slide_height)
for mi, m in enumerate(p.slide_masters):
    for li, lay in enumerate(m.slide_layouts):
        print(f"layout {li} '{lay.name}'")
        for sh in lay.shapes:
            ph = ""
            if sh.is_placeholder:
                ph = f" [ph idx={sh.placeholder_format.idx} type={sh.placeholder_format.type}]"
            pos = "" if sh.left is None else f" L={Emu(sh.left).inches:.2f} T={Emu(sh.top).inches:.2f} W={Emu(sh.width).inches:.2f} H={Emu(sh.height).inches:.2f}"
            print(f"   {sh.shape_type} '{sh.name}'{ph}{pos}")
```

勘察后必须人工确认三件事再动工：
1. 每个版式的占位符 idx（封面标题、目录条目、章节标题、章节编号）；
2. 目录条目行的装饰（编号块/线条）位置，标题文字落点；
3. 主题色与正文字体（读 `ppt/theme/theme1.xml` 的 clrScheme / fontScheme）。
