---
name: bjfu-meetingppt
description: 将组会汇报/工作汇报 Markdown 制作成基于 PPT 母版版式的演示文稿。Use whenever the user wants to 把组会汇报、周报、工作汇报、进展汇报的 markdown 转成 PPT/演示文稿/pptx，或提到 制作汇报PPT、组会PPT、用母版/模板生成PPT、汇报排版 —— 即使用户没有明说"PPT"。默认使用本 skill 封装的母版模板，也支持用户提供自己的模板。
---

# bjfu-meetingppt：组会汇报 Markdown → 母版 PPT

把一份汇报 Markdown（Obsidian 嵌图 `![[xxx.png]]` / 标准链接 `![](x.png)`）制作成成品 PPTX。
所有页面必须从模板母版版式新建（继承母版装饰与页码），不得在白板上自绘装饰。
最终交付**真实的 .pptx**，不是大纲。

## 输入与模板

1. 定位汇报 Markdown。注意：组会汇报使用**同一个固定文件夹**，每周新增一份 md
   （文件名形如 `20261105张佳林组会汇报.md`，含日期与汇报人），文件夹里会积累多周文件：
   - 用户指明了某份 → 用那份；未指明 → 按文件名中的八位日期（YYYYMMDD）取**最新一份**；
   - 无法判断（无日期/同日多份）→ 询问用户，不要猜。
   只处理选定这一份的内容，**不混用其他周的 md**。
   **图片以该 md 中的实际链接为索引定位**（`![[x.png]]` 与 `![](path/x.png)` 均支持），
   用 `collect_md_images(md路径)` 解析（返回与文中出现顺序一致的路径列表）。
   附件夹名不固定（可能叫 `attachments/`、`asserts/` 或其他），且是各周共用的累积目录——
   不要假设夹名、不要扫夹取图，未引用的文件一律忽略；解析为 None 的图片要报告用户，不要静默跳过。
   输出 pptx 与该 md 同名、同目录。
2. 模板：用户指定则用用户的；否则用本 skill 封装的干净母版
   `<skill根目录>/assets/PPT模版.pptx`（结构档案见 `references/template-map.md`，
   **动工前必读**）。换模板时按档案中的勘察代码重新勘察后再改脚本参数。
3. 设计语法参考：`examples/PPT示例_精简版.pptx` 是用户认可的成品示例（渲染图在
   `examples/pages/p01–p19.jpg`），其元素排布与形状使用规范提炼在
   `references/example-design.md`，**定页面计划前必读**并按其执行。

## 第 1 步：读内容、定页面计划

通读 Markdown，按章节映射页面，不发明内容、不改写事实，轻度润字即可：

```
封面(标题=组会工作汇报, 副标题=汇报人：xxx ｜ 日期)   ← 日期/人名取自文件名
目录(各一级标题)
每个项目章节：章节封面(编号+项目名) → 内容页(项目进展 / 后续规划…)
结尾：封面版式放"汇报完毕，谢谢聆听"
```

排版约束（用户验收标准，吸收示例设计语法）：
- **双层标题**：内容页用 `deck.body(章节名, 小节标签)`——左上角青蓝缎带（章节名，如
  "一、消防喷淋头项目"，每页重复强化语境）+ 其下深藏青粗体小节标签（如"1.1 项目进展"）。
- 文字与图片合理占据页面，**不留大面积空白**。
- **每图必注**：图注深藏青粗体 11pt 居中贴图片下缘（`picture(caption=)` 自动贴合）；
  成绩/关键节点图注可用红字。图注只写 Markdown 上下文能支撑的表述，无法确认内容细节的
  图用中性图注并在交付说明中提醒目检。
- **图片两种排法按内容选**：整齐划一（同组设备照片）用 `photo_grid(crop_169=True)`
  等大网格；饱满现场感（调研/现场/多活动照片）用 `pic_notes()` 每行等高自然比例拼贴。
- Markdown 无序列表 → PPT 中保持罗列式（`bullets()` 或 `panel_list()` 面板条）。
- **红字关键词**：要点与底部成果条中，关键成果/数字用红字强调（每页 1–3 处，勿滥用）；
  每页可加 `footer_note()` 底部成果/总结条收尾。
- 表格默认 `table()` 青底白字表头；信息密集的成果/课程类表格可用 `style="grid"`
  （示例网格风：白底黑粗体表头+隔行淡蓝）。
- 技术路线/工作流程页用 `flow_cards()`（彩色标题条步骤卡+块箭头）+ `frame()` 细线框。
- 页面内容若只有一张超宽图（如甘特/纪要表），整块（引言+图+图注）在内容区垂直居中。

## 第 2 步：写内容脚本并构建

在 Markdown 同级目录建当周工作夹 `_ppt_build/<YYYYMMDD>/`（文件夹固定复用，按周隔离，
避免覆盖往周脚本与报告），写 `build_ppt.py`：import 本 skill 的 `scripts/ppt_kit.py`
（含全部踩坑修复：MPO 转码、中文 a:ea 字体、项目符号 XML、清页 drop_rel、等比放图）。
示例：

```python
import os, sys
SKILL = r"<skill根目录>"
sys.path.insert(0, os.path.join(SKILL, "scripts"))
from ppt_kit import Deck, collect_md_images, RED

refs, paths = collect_md_images("20261105张佳林组会汇报.md")  # 按 md 实际链接定位图片
# paths 与 md 中出现顺序一致，缺失项为 None（需报告用户）
deck = Deck(os.path.join(SKILL, "assets", "PPT模版.pptx"), work_dir="_ppt_build/20261105")
deck.cover("组会工作汇报", "汇报人：张佳林 ｜ 2026年11月5日")
deck.toc(["项目一", "项目二", "其他工作"])
deck.section(1, "项目一")
s = deck.body("一、项目一", "1.1 项目进展：第一阶段原理验证完成")
cy = s.content_y                                   # 缎带+小节标签之下的内容起始 y
s.bullets(["已完成方案原理验证", "末端执行结构架构确认", "电控系统开发完毕"], 0.45, cy + 0.02, 12.43)
s.picture(paths[0], 0.45, cy + 1.30, 6.55, 3.47, border=True, caption="现场验证")
s.table([("人员", "分工"), ("张三", "机械设计")], 7.30, cy + 1.30, 5.58, [1.75, 3.83])
s.footer_note([("第一阶段验证完成，", None), ("电控系统已就绪", RED), ("，转入现场测试。", None)])
deck.section(2, "项目二")
s = deck.body("二、项目二", "2.1 技术路线")
s.flow_cards([dict(title="1.语义输入", text="开放语义导航"),
              dict(title="2.Agent规划", pic=paths[1]),
              dict(title="3.小车执行", arrow="差速底盘"),
              dict(title="4.Demo验证", text="实车验证")], 0.70, cy + 0.35, 11.9, 3.8)
s2 = deck.body("二、项目二", "2.2 后续计划"); s2.panel_list([...5条...])
s3 = deck.body("二、项目二", "2.3 现场调研"); s3.pic_notes([...图+图注...], 0.45, cy, 12.43, row_h=2.6)
deck.cover("汇报完毕，谢谢聆听", "汇报人：张佳林 ｜ 2026年11月5日", title_size=40)
deck.save("20261105张佳林组会汇报.pptx")   # 输出名与 md 同名
```

输出文件放在 Markdown 同级目录、与 Markdown 同名。

## 第 3 步：质检与修正（必做一轮）

1. `python <skill>/scripts/qa_deck.py check <pptx>` —— 越界/重叠必须为 0。
2. 渲染：`from ppt_kit import export_pngs; export_pngs(pptx, "_render/deck")`
   （Windows PowerPoint COM；无 PowerPoint 时跳过渲染，只做几何检查并向用户说明）。
3. `python <skill>/scripts/qa_deck.py pixels _render/deck` —— 覆盖率异常低/空白格≥5 的页
   记为疑点（注意：白底文档类图片会误报"空白"，先看渲染图再定）。
4. 把渲染图按页分派给视觉评审代理（`presentations:visual-judge`），逐页要
   `{"page":n,"pass":bool,"issues":[...]}`；不合格页修复后重渲染、复审。
   常见修法：空页→内容块垂直居中或放大图片；照片组不等大→`photo_grid(crop_169=True)`；
   卡片头重脚轻→压缩卡高并让内容垂直居中。

## 第 4 步：交付

- 成品 pptx + `_ppt_build/qa_report.md`（页面结构表、质检过程、已知事项）。
- 交付说明中必须写明：① 页面结构一览；② 图片做了 16:9 自动裁剪的请用户目检主体完整；
  ③ **发现疑似放错章节的图片时不要擅自移动**，按原位置保留、图注中性化，并在说明中显著
  提请用户确认是否移动。
