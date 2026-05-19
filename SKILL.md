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

**定位脚本路径**：脚本随技能一同安装在技能目录的 `scripts/` 子目录下。运行以下命令找到实际路径：
```powershell
Get-ChildItem -Path "$HOME\.claude\skills" -Recurse -Filter "Convert-SvgToShapes.ps1" | Select-Object FullName
```
找到后记录绝对路径，后续所有命令均替换占位符 `<SKILL_SCRIPTS>`。

## 路径选择

```
用户有品牌模板（.pptx）？
  ├─ 是  →  路径 B：编辑品牌模板
  └─ 否  →  路径 A：从零生成演示文稿
```

## 路径 A：从零生成演示文稿

### 第一步：明确需求（快速决策）

以下是**无需询问可直接假设的默认值**：

| 信息 | 默认值 |
|------|--------|
| 幻灯片数量 | 6 张（标题页 + 4 内容页 + 结尾页） |
| 语言 | 用户对话语言 |
| 风格 | 商务简约 |
| 输出路径 | 当前工作目录 |

**主题、受众、核心信息**是唯一需要确认的信息——若缺失，一次性提问。

### 第二步：规划结构与配色方案

先定好两件事再动笔：幻灯片分工、颜色体系。配色一旦确定，整套幻灯片不再更改。

**幻灯片分工**：

| # | 类型 | 内容定位 |
|---|------|----------|
| 1 | 标题页 | 演示标题 + 副标题 + 日期/署名 |
| 2–N-1 | 内容页 | 每页聚焦一个论点，不超过 5 个要点 |
| N | 结尾页 | 结论 / 下一步 / 联系方式 |

**配色方案（三选一，或根据用户偏好自定义）**：

| 方案 | 主色（背景/标题） | 强调色（高亮/图标） | 卡片背景 |
|------|-----------------|---------------------|----------|
| 活力蓝商务 | `#1D4ED8` | `#60A5FA` | `#EFF6FF` |
| 墨绿科技 | `#1A3C34` | `#2ECC71` | `#F0F4F2` |
| 暖橙活力 | `#7B2D00` | `#E67E22` | `#FAFAFA` |

每张内容页选一种视觉结构——读 `references/svg-design.md` 判断最适合的结构。**选择顺序：优先匹配"用户布局库"（UL-1 至 UL-5，来自用户实际使用的版式）；无合适匹配时再使用通用结构（并列、总分、流程、循环、金字塔、对比、聚焦）**。

### 第三步：逐张生成并立即验证 SVG

**核心约束**：脚本以 letterbox 模式插入 SVG（保持宽高比）——若 SVG 本身有大量空白，缩放后会出现双重空白，内容显得很小。**边距通过坐标控制（x: 40–920，y: 20–520），不要依赖 viewBox 外部的留白**。

**处理节奏：逐张完成——生成 → 立即自检 → 发现问题立即修正 → 再进入下一张。**

每张幻灯片的 SVG 从以下骨架开始：

```xml
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 960 540">
  <!-- 第一个元素：铺满背景，必须存在 -->
  <rect width="960" height="540" fill="#主色或背景色"/>
  <!-- 标题区：顶部 0–80px -->
  <!-- 内容区：y: 80–520，x: 40–920 -->
</svg>
```

生成后立即执行以下自检（每项不通过则修正再继续）：

**文字越界**（逐个 `<text>` 元素心算）
- 中文每字宽度 ≈ `font-size × 1.0` px；英文小写每字母 ≈ `font-size × 0.55` px
- `text-anchor="start"`：右边界 = x + 总宽 ≤ 920
- `text-anchor="middle"`：左边界 = x − 总宽/2 ≥ 40，右边界 = x + 总宽/2 ≤ 920
- `text-anchor="end"`：左边界 = x − 总宽 ≥ 40
- y 基线 + 5px（descender）≤ 520；y 基线 − font-size ≥ 20
- 单行超过 20 个汉字（或 36 个英文字母）必须手动拆行（行距 = font-size × 1.4）

```
□ 无单行文字估算后超出 x: 40–920
□ 无文字 y 超出 20–520
□ 长文本已手动换行
```

**图形越界**（逐个形状核对边界）

```
□ <rect x y width height>：x ≥ 0，y ≥ 0，x+width ≤ 960，y+height ≤ 540
    有 stroke 时：x − stroke-width/2 ≥ 0，x+width + stroke-width/2 ≤ 960（y 方向同理）
□ <circle cx cy r>：cx−r ≥ 0，cx+r ≤ 960，cy−r ≥ 0，cy+r ≤ 540
□ <path> / <polygon> 所有坐标点均在 x: 0–960，y: 0–540
```

**布局质量**

```
□ 第一个元素是铺满 viewBox 的背景 <rect width="960" height="540" fill="..."/>
□ 所有内容距 viewBox 边缘 ≥ 40px（x: 40–920，y: 20–520）
□ 同类元素对齐（多列卡片的 y/height 完全一致）
□ 颜色使用本套已确定的色板（主色/强调色跨幻灯片一致）
□ 标题字号 ≥ 28px，正文字号 ≥ 16px
□ 单张文字行数 ≤ 12 行
□ 所选视觉结构与内容关系类型匹配
```

**技术兼容性**（违反则静默栅格化，失去可编辑性）

```
□ 无 <filter>（模糊/投影）
□ 无 <clipPath> / <mask>
□ 无 <image> 嵌入位图（改用形状组合）
□ 无 <style> 块（改用内联 style= 属性）
□ 字体为系统字体（微软雅黑、Arial、Calibri）
□ 渐变只用 linearGradient，色标 ≤ 3 个
```

### 第四步：运行 Convert-SvgToShapes.ps1

```powershell
# 单个 SVG
pwsh -File "<SKILL_SCRIPTS>\Convert-SvgToShapes.ps1" `
     -SvgPath "C:\slides\slide1.svg" `
     -OutputPath "C:\out\deck.pptx" -Force

# 一个目录里的所有 SVG（按文件名排序，每个占一张幻灯片）
pwsh -File "<SKILL_SCRIPTS>\Convert-SvgToShapes.ps1" `
     -SvgPath "C:\slides\" `
     -OutputPath "C:\out\deck.pptx" -Force
```

**常见问题**：
- 脚本挂起无响应 → 确认加了 `-Force`，检查 PowerPoint 是否弹出了未关闭的对话框
- 某张幻灯片仍显示"转换为形状"选项 → 该 SVG 含不支持特性，查看脚本 stderr 输出定位

### 第五步：向用户报告结果

脚本成功后告知用户：
1. 输出文件的完整路径
2. 共生成几张幻灯片
3. 验证方法：右键点击形状，若出现"编辑顶点"则转换成功；若仍显示"转换为形状"则需排查该 SVG

## 路径 B：编辑品牌模板

### 第一步：检查模板结构（必做）

在生成任何内容之前，先了解模板结构，避免索引出错。运行：

```powershell
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

记录：封面在哪张、目录在哪张、内容模板在哪张（供 EXPAND 时复制）、共几张幻灯片。

### 第二步：选择编辑模式

```
封面 / 目录页改文字              →  TEXT 模式
替换某张已有幻灯片的内容          →  INSERT 模式
复制内容模板 + 新增多张内容页      →  EXPAND 模式
涉及 2 种及以上操作              →  MANIFEST 模式（推荐）
```

**操作步骤达到 3 个或以上时，始终用 MANIFEST 模式**——一个 JSON 文件描述全部操作，一次调用完成，避免反复启动 PowerPoint。

### 第三步：生成内容页 SVG（INSERT / EXPAND 模式需要）

EXPAND 模式下，SVG **只负责内容区的插图**——不要在 SVG 里重绘模板的标题栏、Logo、页脚、背景，这些来自复制的模板幻灯片。

SVG 设计为填满自身 viewBox，脚本会按比例缩放后居中放入内容区。逐张生成并执行路径 A 第三步的自检清单（技术约束完全相同）。

### 第四步：运行 Edit-ExistingPptx.ps1

**TEXT 模式**（只改文字）：

```powershell
pwsh -File "<SKILL_SCRIPTS>\Edit-ExistingPptx.ps1" `
     -TemplatePath "C:\branded.pptx" -OutputPath "C:\out.pptx" -Force `
     -SlideTexts @{
         1 = @{ Title = '2026 战略报告'; Subtitle = '董事会汇报'; Date = '2026-05' }
         2 = @{ Title = '目录'; Body = @('市场概述','竞争格局','战略规划') }
     }
```

**MANIFEST 模式**（推荐用于复杂任务）：

```powershell
pwsh -File "<SKILL_SCRIPTS>\Edit-ExistingPptx.ps1" `
     -TemplatePath "C:\branded.pptx" -Manifest "C:\deck.json" `
     -OutputPath "C:\out.pptx" -Force
```

Manifest 结构见下方；SVG 路径相对于 Manifest 文件所在目录。

## SVG 技术约束快查表

PowerPoint 的 SVGEdit 转换器只支持 SVG 的一个子集。不符合约束的元素会被**静默栅格化**（变成图片，失去可编辑性）。

**可以用**
- 基本形状：`<path>` `<rect>` `<circle>` `<ellipse>` `<line>` `<polygon>` `<polyline>`
- 文字：`<text>` 配合内联 `style` 或属性（`font-family` `font-size` `fill` 等）
- 纯色填充与描边（十六进制 / rgb）
- 简单 `linearGradient`（≤3 色标）
- `<g>` 分组与基本变换（`translate` `scale` `rotate`）

**不能用**

| 特性 | 结果 |
|------|------|
| `<filter>`（模糊、投影） | 整个元素栅格化为位图 |
| `<clipPath>` `<mask>` | 通常被静默丢弃 |
| `<image href="…">` 嵌入位图 | 变为图片，不可编辑 |
| `radialGradient`、网格渐变 | 通常不支持 |
| `<style>` 块内 CSS | 改用内联 `style=` 属性 |
| Web 字体 / `@font-face` | 改用系统字体（微软雅黑、Arial、Calibri）|
| `<foreignObject>` | 完全不支持 |

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

## 自动化调用要点

1. **始终用绝对路径**——COM 对相对路径处理不稳定
2. **始终加 `-Force`**——脚本在非交互式环境缺少此参数会挂起
3. **优先用 `pwsh` 而非 `powershell.exe`**——`pwsh` 默认 UTF-8，中文字符更稳定
4. **中文内容 + 多张幻灯片时，写 Manifest JSON 文件再调用 `-Manifest`**——完全规避 Shell 引号转义问题
5. **PowerPoint 窗口在转换期间必须可见**——脚本会自动设置 `Visible = True`
6. **转换前关闭其他演示文稿**——打开的文件可能干扰选择状态

## 已知局限

- **每张幻灯片只能插入一个 SVG**（单页多 SVG 布局不在支持范围）
- **不支持无头 / CI 环境**（ExecuteMso 需要可见窗口）
- **仅支持 Windows + PowerPoint 2016 build 1712+**
- SVG → 形状的还原效果取决于 PowerPoint 内置转换器；复杂 SVG 可能部分栅格化

**SVG 幻灯片设计参考**：`references/svg-design.md`（常见布局模板、配色指南、字体建议）
