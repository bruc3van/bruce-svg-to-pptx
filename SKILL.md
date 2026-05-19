---
name: bruce-svg-to-pptx
description: 当用户想用 AI 生成演示文稿、将 SVG 转为 PowerPoint 可编辑形状，或在现有品牌 PPTX 模板中填入 AI 内容时使用。触发场景包括：「帮我做一个关于 X 的 PPT」「生成一份演示文稿」「SVG 转 PPT 可编辑形状」「把这些 SVG 导入 PowerPoint」「基于模板生成 PPT」「修改封面/目录文字」「批量转换 SVG 图标」，或任何涉及 AI 生成幻灯片内容并输出为可编辑 PowerPoint 的需求。只要用户提到"PPT""演示文稿""幻灯片"并期待可编辑输出，就应触发本技能。仅支持 Windows。
---

# SVG → PowerPoint 可编辑形状（Windows）

## 核心原理

AI 生成 SVG → PowerPoint 用 `ExecuteMso("SVGEdit")` 将其转换为原生 DrawingML 形状 → 结果与手动绘制完全一致：可取消组合、可改色、可添加动画。

本技能附带两个脚本（`scripts/` 目录）：
- **`Convert-SvgToShapes.ps1`** — 从 SVG 新建演示文稿
- **`Edit-ExistingPptx.ps1`** — 在现有品牌模板中编辑文字 / 插入 SVG

---

## 路径选择

```
用户有品牌模板（.pptx）？
  ├─ 是  →  路径 B：编辑品牌模板
  └─ 否  →  路径 A：从零生成演示文稿
```

---

## 路径 A：从零生成演示文稿

### 第一步：与用户确认需求

在开始生成前，先明确以下几点（可一次性询问，避免来回）：

- **主题 / 受众**：演示给谁看？核心信息是什么？
- **幻灯片数量**：默认 5–8 张（标题页 + 内容 + 结尾）
- **语言**：中文 / 英文 / 混合
- **风格**：商务正式 / 简约现代 / 科技感 / 其他
- **输出路径**：保存到哪里？（告知用户默认当前目录）

如果用户已给出足够信息（主题明确、数量合理），可跳过询问直接执行。

### 第二步：规划幻灯片结构

给每张幻灯片定一个明确的职责，典型结构：

| # | 类型 | 内容定位 |
|---|------|----------|
| 1 | 标题页 | 演示标题 + 副标题（可含日期/署名） |
| 2–N-1 | 内容页 | 每页聚焦一个论点 / 数据 / 流程 |
| N | 结尾页 | 结论 / 下一步 / 联系方式 |

**每张内容页选一种布局**（参见 `references/svg-design.md`）：要点列表、三列对比、流程步骤、数据聚焦、图文混排。**不要让一张幻灯片承载超过 5 个要点**。

### 第三步：生成每张幻灯片的 SVG

为每张幻灯片单独生成一个 `.svg` 文件。**严格遵守下方 SVG 技术约束**，否则 PowerPoint 无法将其转换为可编辑形状（会降级为图片）。

**关键：SVG 必须铺满 viewBox——不要留白**

```xml
<!-- 始终从一个铺满整个 viewBox 的背景矩形开始 -->
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 960 540">
  <!-- 第一个元素：背景 -->
  <rect width="960" height="540" fill="#1E3A5F"/>
  <!-- 之后才是其他内容 -->
  ...
</svg>
```

脚本在插入时会**保持宽高比**（letterbox 模式）。若 SVG 本身有大量空白边距，缩放后就会出现双重空白，导致内容看起来很小。**将内容设计为填满整个 viewBox，边距在 SVG 内部用 padding 控制**（而非依赖外部缩放来留边）。

设计原则（详见 `references/svg-design.md`）：

- **viewBox 用 `"0 0 960 540"`**（16:9）或 `"0 0 720 540"`（4:3）
- 第一个元素必须是铺满 viewBox 的背景 `<rect>`
- 每张幻灯片留标题区（顶部 60–80px 内）和内容区（剩余 460–480px）
- 颜色用 2–3 色主调 + 白/灰中性色；避免超过 5 种颜色
- 正文字号 ≥ 18px，标题 ≥ 28px；使用微软雅黑 / Arial / Calibri
- SVG 里包含完整内容（标题 + 正文），不依赖 PowerPoint 占位符

### 第三步（续）：SVG 自检清单（生成后逐张自我审查，发现问题立即修正）

**一、文字越界检查（逐个 `<text>` 元素心算）**

SVG 不会自动换行，越界文字会直接溢出幻灯片或被截断——这是最常见的问题。

估算规则：
- 中文每字宽度 ≈ `font-size × 1.0` px
- 英文小写每字母 ≈ `font-size × 0.55` px，大写 ≈ `0.7` px
- `text-anchor="start"`：右边界 = x + 估算总宽，需 ≤ 920
- `text-anchor="middle"`：左边界 = x − 总宽/2 需 ≥ 40，右边界 = x + 总宽/2 需 ≤ 920
- `text-anchor="end"`：左边界 = x − 估算总宽，需 ≥ 40
- y 基线 + 约 5px（descender）≤ 520；y 基线 − font-size ≥ 20

```
□ 没有任何单行文字估算后超出 x: 40–920 范围
□ 没有任何文字的 y 坐标超出 20–520 范围
□ 单行超过 20 个汉字（或 36 个英文字母）的文本已手动拆为多行
  （行间距 = font-size × 1.4，如 16px 字号行距约 22px）
```

**二、图形越界检查（逐个形状元素核对边界）**

```
□ <rect x y width height>：x ≥ 0，y ≥ 0，x+width ≤ 960，y+height ≤ 540
    有 stroke 时：x − stroke-width/2 ≥ 0，x+width + stroke-width/2 ≤ 960（同理 y 方向）
□ <circle cx cy r>：cx−r ≥ 0，cx+r ≤ 960，cy−r ≥ 0，cy+r ≤ 540
□ <ellipse cx cy rx ry>：cx−rx ≥ 0，cx+rx ≤ 960，cy−ry ≥ 0，cy+ry ≤ 540
□ <line x1 y1 x2 y2>：x1,x2 在 0–960，y1,y2 在 0–540
□ <polygon> / <polyline> 所有顶点：x 在 0–960，y 在 0–540
□ <path>：所有 M/L/C/Q 命令的坐标点均在 viewBox 内
```

**三、布局质量检查**

```
□ 第一个元素是铺满 viewBox 的背景 <rect width="960" height="540" fill="..."/>
□ 所有内容距 viewBox 边缘 ≥ 40px（左右 x: 40–920，上下 y: 20–520）
□ 同类元素对齐：多列卡片的 y/height 完全一致；多行文字的 x 坐标一致
□ 视觉重心均衡，无一侧大片空白
□ 所选视觉结构（并列/流程/金字塔等）与内容的本质关系匹配
```

**四、可读性检查**

```
□ 标题字号 ≥ 28px，正文字号 ≥ 16px，注释字号 ≥ 13px
□ 浅色背景上文字用深色（#1E3A5F 或 #2D3748），深色背景上用白色（#FFFFFF）
□ 单张幻灯片总文字行数 ≤ 12 行，避免信息过载
□ 整套演示文稿颜色体系统一（主色/强调色/背景色跨幻灯片一致）
```

**五、技术兼容性检查（PowerPoint SVGEdit 约束）**

```
□ 无 <filter>（模糊/投影）——会导致整个元素栅格化为位图
□ 无 <clipPath> / <mask>——会被静默丢弃
□ 无 <image> 嵌入位图——改用形状组合模拟
□ 无 <style> 块——所有样式改为内联 style= 属性或元素属性
□ 字体为系统字体（微软雅黑、Arial、Calibri 等），无 Web 字体
□ 渐变只用 linearGradient，色标 ≤ 3 个
```

全部通过后再执行第四步，节省反复修改的时间。

### 第四步：运行 Convert-SvgToShapes.ps1

```powershell
# 单个 SVG
pwsh -File "C:\path\scripts\Convert-SvgToShapes.ps1" `
     -SvgPath "C:\slides\slide1.svg" `
     -OutputPath "C:\out\deck.pptx" -Force

# 一个文件夹里的所有 SVG（按文件名排序，每个占一张幻灯片）
pwsh -File "C:\path\scripts\Convert-SvgToShapes.ps1" `
     -SvgPath "C:\slides\" `
     -OutputPath "C:\out\deck.pptx" -Force
```

**始终用绝对路径**；**始终加 `-Force`**（脚本在非交互环境中若缺少此参数会挂起等待确认）。

---

## 路径 B：编辑品牌模板

### 第一步：检查模板（必做）

在生成任何 SVG 之前，先了解模板结构，避免索引出错。

```powershell
# 用这段 PowerShell 快速列出所有幻灯片
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1
$pres = $ppt.Presentations.Open("C:\path\branded.pptx", 0, 0, -1)
1..$pres.Slides.Count | ForEach-Object {
    $s = $pres.Slides.Item($_)
    $title = try { $s.Shapes.Title.TextFrame.TextRange.Text } catch { "(无标题)" }
    Write-Host "$_ : $title  [布局: $($s.CustomLayout.Name)]"
}
$pres.Close(); $ppt.Quit()
```

了解后确认：
- 哪张是封面？哪张是目录？哪张是"内容模板"（EXPAND 时复制它）？
- 内容模板幻灯片是否有 body 占位符？（有占位符 = 内容区自动检测更准）
- 总共几张？用户要在哪里插入新内容？

### 第二步：规划编辑方案

```
封面 / 标题页 / 目录页   →  TEXT 模式（只改文字）
替换某张已有幻灯片        →  INSERT 模式
复制内容模板 + 插入 N 张  →  EXPAND 模式
以上组合 / 完整演示文稿   →  MANIFEST 模式（推荐）
```

**只要同时涉及"改文字"和"加内容页"，优先用 MANIFEST 模式**——一个 JSON 描述所有操作，一次调用完成，避免反复启动 PowerPoint。

### 第三步：生成内容页 SVG（仅 INSERT / EXPAND 时需要）

EXPAND 模式下，SVG **只负责内容区的插图**——不要在 SVG 里重绘模板的标题栏、Logo、页脚、背景，这些来自复制的模板幻灯片。

SVG 设计为填满自身 viewBox，脚本会按比例缩放后居中放入内容区。

### 第四步：运行 Edit-ExistingPptx.ps1

**TEXT 模式**（只改文字）：
```powershell
pwsh -File "C:\path\scripts\Edit-ExistingPptx.ps1" `
     -TemplatePath "C:\branded.pptx" -OutputPath "C:\out.pptx" -Force `
     -SlideTexts @{
         1 = @{ Title = '2026 战略报告'; Subtitle = '董事会汇报'; Date = '2026-05' }
         2 = @{ Title = '目录'; Body = @('市场概述','竞争格局','战略规划') }
     }
```

**MANIFEST 模式**（推荐用于复杂任务）：
```powershell
pwsh -File "C:\path\scripts\Edit-ExistingPptx.ps1" `
     -TemplatePath "C:\branded.pptx" -Manifest "C:\deck.json" `
     -OutputPath "C:\out.pptx" -Force
```

Manifest 结构见下方参考；SVG 路径相对于 Manifest 文件所在目录。

---

## SVG 技术约束

PowerPoint 的 SVGEdit 转换器只支持 SVG 的一个子集。**不符合约束的元素会被静默栅格化**（变成图片，失去可编辑性）。

**可以用**

- 基本形状：`<path>` `<rect>` `<circle>` `<ellipse>` `<line>` `<polygon>` `<polyline>`
- 文字：`<text>` 配合内联 `style` 或属性（`font-family` `font-size` `fill` 等）
- 纯色填充与描边（十六进制 / rgb）
- 简单 `linearGradient`（≤3 色标）
- `<g>` 分组与基本变换（`translate` `scale` `rotate`）

**不能用（会被栅格化或丢弃）**

| 特性 | 结果 |
|------|------|
| `<filter>`（模糊、投影） | 整个元素栅格化为位图 |
| `<clipPath>` `<mask>` | 通常被静默丢弃 |
| `<image href="…">` 嵌入位图 | 变为图片，不可编辑 |
| `radialGradient`、网格渐变 | 通常不支持 |
| `<style>` 块内 CSS | 改用内联 `style=` 属性 |
| Web 字体 / `@font-face` | 改用系统字体（微软雅黑、Arial、Calibri）|
| `<foreignObject>` | 完全不支持 |

---

## 脚本参考

### Convert-SvgToShapes.ps1 参数

| 参数 | 必填 | 说明 |
|------|------|------|
| `-SvgPath` | ✓ | `.svg` 文件或目录（支持数组） |
| `-OutputPath` | ✓ | 目标 `.pptx`，不存在时自动创建 |
| `-Append` | | 追加到已有文件而非覆盖 |
| `-Force` | | 跳过覆盖确认（自动化调用必加） |

### Edit-ExistingPptx.ps1 参数

| 参数 | 说明 |
|------|------|
| `-TemplatePath` | 要编辑的已有 `.pptx`（必填） |
| `-OutputPath` | 输出路径，默认原地编辑（自动备份） |
| `-Manifest` | JSON manifest 文件路径（MANIFEST 模式） |
| `-SvgPath` | SVG 文件或目录（INSERT / EXPAND 模式） |
| `-TargetSlide` | INSERT 模式：目标幻灯片索引（1 起） |
| `-ContentSlide` | EXPAND 模式：用作模板的幻灯片（默认末张） |
| `-InsertAfterSlide` | EXPAND 模式：新幻灯片插入在此之后 |
| `-ContentZone` | 覆盖内容区域：`"Left,Top,Width,Height"`（单位：点） |
| `-SlideTexts` | 文字编辑哈希表（见下） |
| `-ClearContent` | 插入前清除内容区内的非结构性形状 |
| `-NoBackup` | 跳过自动创建 `.bak.pptx` |
| `-Force` | 跳过覆盖确认（自动化必加） |

**SlideTexts 条目支持的键**：`Title` `Subtitle` `Body`（字符串或数组）`Date` `Footer` `ShapeTexts`（按名称匹配文本框）`Replacements`（全局查找替换）。

### Manifest 结构

```json
{
  "edits": [
    { "type": "text",   "slide": 1,
      "title": "...", "subtitle": "...", "body": ["..."],
      "shapeTexts": [{ "shapeName": "Tagline", "text": "..." }],
      "replacements": [{ "find": "{{year}}", "replace": "2026" }] },

    { "type": "expand", "templateSlide": 3, "insertAfter": 2,
      "clearContent": true,
      "items": [
        { "svg": "slides/s1.svg", "title": "第一章" },
        { "svg": "slides/s2.svg", "title": "第二章" }
      ] },

    { "type": "insert", "slide": 9, "svg": "slides/closing.svg",
      "clearContent": true, "title": "结语" }
  ]
}
```

---

## 自动化调用要点

1. **始终用绝对路径**——COM 对相对路径处理不稳定。
2. **始终加 `-Force`**——脚本在非交互式环境缺少此参数会挂起。
3. **优先用 `pwsh` 而非 `powershell.exe`**——`pwsh` 默认 UTF-8，中文字符更稳定。
4. **中文内容 + 多张幻灯片时，写 Manifest JSON 文件再调用 `-Manifest`**——完全规避 Shell 引号转义问题。
5. **PowerPoint 窗口在转换期间必须可见**——脚本会自动设置 `Visible = True`。
6. **转换前关闭其他演示文稿**——脚本会发出警告但不中止；打开的文件可能干扰选择状态。

---

## 验证结果

打开生成的 `.pptx`，右键单击插入的形状：
- **显示"编辑顶点"** 或可取消组合 → 转换成功，是真正的可编辑形状
- **仍显示"转换为形状"** → SVG 包含不支持的特性；查看脚本 stderr 定位原因

---

## 已知局限

- **每张幻灯片只能插入一个 SVG**（单页多 SVG 布局不在支持范围）
- **不支持无头 / CI 环境**（ExecuteMso 需要可见窗口）
- **仅支持 Windows + PowerPoint 2016 build 1712+**
- SVG → 形状的还原效果取决于 PowerPoint 内置转换器；复杂 SVG 可能部分栅格化

---

**SVG 幻灯片设计参考**：`references/svg-design.md`（常见布局模板、配色指南、字体建议）
