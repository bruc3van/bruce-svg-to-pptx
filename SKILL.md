---
name: bruce-svg-to-pptx
description: 当用户想在 Windows 上将 SVG 文件转换为 PowerPoint 可编辑原生形状，或将 SVG 内容插入现有品牌 PPTX 模板（保留封面/目录/布局页，同时编辑文字或新增内容页）时使用。触发场景包括："SVG → 形状"、"SVG 转 PPT 可编辑形状"、"插入到现有 PPT"、"基于模板生成"、批量转换 SVG 图标、为演示稿扩充内容页、仅修改封面/目录页文字，或任何希望让 SVG 在 PowerPoint 中变为可编辑形状的需求。仅支持 Windows。
---

# SVG 转 PowerPoint 可编辑形状（Windows）

## 核心思路

本技能的核心理念：由 LLM 将完整内容页**以 SVG 形式**布局（图形 + 文字），PowerPoint 导入该 SVG 后，通过 `CommandBars.ExecuteMso("SVGEdit")` 将其转换为原生可编辑形状。最终结果是真正的 PowerPoint DrawingML 形状组合——可取消组合、可重新着色、可添加动画、完全可编辑。

当用户持有**品牌模板**（公司封面、目录、内容布局）时，目标转变为：保留现有品牌页面，不同页面类型走不同处理路径：

```
封面 / 标题页 / 目录页    →  TEXT 模式
                          仅替换占位符里的文字（标题、副标题、目录条目…）
                          不插入 SVG。整页布局来自模板。

正文 / 内容页             →  EXPAND 模式
                          复制模板的"内容页"做底稿，保留它的标题栏、背景、
                          页脚等品牌元素；在中间的安全区里嵌入一张较小的
                          LLM 生成 SVG，再把它转换成可编辑形状。

完整一份演示稿            →  MANIFEST 模式
                          一份 JSON 描述每一页的编辑（封面文字 + 目录文字
                          + N 张正文页插图），单次调用一次性产出。
```

## 本技能的功能

通过 COM 在 Windows 上驱动 Microsoft PowerPoint，执行用户手动操作时的等效步骤：

1. 将每个 SVG 作为 `msoGraphic` 形状插入幻灯片。
2. 选中后调用 `CommandBars.ExecuteMso("SVGEdit")` —— 即 PowerPoint 内置的"转换为形状"功能区命令。VBA 没有直接方法；此 Mso 命令是 Microsoft 暴露的唯一编程接口。
3. 保存演示文稿。

本技能附带**两个脚本**：

| 脚本 | 用途 |
|---|---|
| `scripts/Convert-SvgToShapes.ps1` | 从一个或多个 SVG **新建** .pptx（每个 SVG 占一张幻灯片，铺满全页）。 |
| `scripts/Edit-ExistingPptx.ps1`   | 在**现有**品牌模板中编辑文字和/或插入 SVG 内容页。支持四种模式：TEXT、INSERT、EXPAND、MANIFEST。 |

## 决策树 —— 选择正确的调用方式

```
用户需求                                                       → 使用
────────────────────────────────────────────────────────────────────────
"从这些 SVG 生成一份全新演示文稿"                              → Convert-SvgToShapes.ps1
"批量将这些 SVG 图标转换到一份演示文稿"                        → Convert-SvgToShapes.ps1

"我有品牌模板，只需修改封面/目录/标题页的文字"                 → Edit-ExistingPptx.ps1（TEXT 模式）

"我有品牌模板，用这个 SVG 替换第 N 张幻灯片的内容"            → Edit-ExistingPptx.ps1（INSERT 模式）

"我有品牌模板，将内容布局复制 N 次，每次插入一个 SVG"         → Edit-ExistingPptx.ps1（EXPAND 模式）

"从品牌模板生成完整演示稿——封面文字 + 目录文字 + N 张内容页
 SVG——一次搞定"                                               → Edit-ExistingPptx.ps1（MANIFEST 模式）
```

## SVG 生成规范（面向 LLM）

PowerPoint 内置的 SVG → 形状转换器是瓶颈所在。要获得干净、完全可编辑的输出，生成的 SVG **必须**限定在其支持的子集内：

**可安全使用**

- `<path>`、`<rect>`、`<circle>`、`<ellipse>`、`<line>`、`<polygon>`、`<polyline>`
- 带有 `font-family`、`font-size`、`font-weight`、`fill` 属性的 `<text>`
- 纯色 `fill` 和 `stroke`（十六进制、rgb）
- 简单 `linearGradient`（2–3 个色标）
- 带基本变换（`translate`、`scale`、`rotate`）的 `<g>` 分组

**避免使用（会被栅格化、静默丢弃或转换失败）**

- `<filter>`（投影、模糊等）—— 会被栅格化为位图
- `<clipPath>`、`<mask>` —— 通常被静默丢弃
- `<foreignObject>` —— 完全不支持
- `<image href="…">`（嵌入位图）—— 失去可编辑性，结果变为图片而非形状
- `radialGradient`、网格渐变、超过 3 个色标的渐变
- `<style>` 块中的 CSS —— 只用内联 `style=` 或属性
- Web 字体 / `@font-face` —— PowerPoint 无法替换；改用系统字体（Arial、Calibri、思源黑体、微软雅黑）或将文字转为路径
- `pattern` 填充 —— 经常丢失

**内容页 SVG 的布局规则（EXPAND 模式）**

SVG 被插入到**安全内容区**，而非整张幻灯片。默认该区域为内容模板幻灯片的 body 占位符范围；若无 body 占位符，则为标题栏下方带边距的区域。**请将 SVG 设计为填满自身 viewBox，不要假设具体的幻灯片尺寸** —— 脚本会保持宽高比（信箱格式）并将 SVG 居中于该区域。对于品牌模板，**不要**在 SVG 内部重新绘制模板的标题栏、页脚、Logo 或背景；这些元素由复制的 PowerPoint 模板幻灯片提供。SVG 应只包含属于安全内容区的可编辑内容插图、图表或文字块。

## 前置条件

- Windows 10 / 11
- Microsoft PowerPoint 2016 版本 1712+（任何当前 Microsoft 365 / Office 2019 / 2021 / 2024 均满足）
- PowerShell 5.1（Windows 内置）或 PowerShell 7（`pwsh`，**推荐**，默认 UTF-8，中文字符更稳定）

若未安装 PowerPoint，脚本会输出清晰的错误信息并退出。

## 脚本一 —— `Convert-SvgToShapes.ps1`

创建一份新 .pptx，每个 SVG 占一张幻灯片（铺满全页，保持宽高比）；也可追加到已有演示文稿。

| 参数 | 必填 | 说明 |
|---|---|---|
| `-SvgPath`    | ✓ | `.svg` 文件或目录，支持数组。 |
| `-OutputPath` | ✓ | 目标 `.pptx`，不存在时自动创建。 |
| `-Append`     |   | 追加到已有 `-OutputPath`，而非覆盖。 |
| `-Force`      |   | 跳过覆盖确认。 |

```powershell
# 单个 SVG → 新建演示文稿
pwsh -File scripts\Convert-SvgToShapes.ps1 -SvgPath .\icon.svg -OutputPath .\out.pptx

# 文件夹中的 SVG，每个占一张幻灯片
pwsh -File scripts\Convert-SvgToShapes.ps1 -SvgPath .\icons\ -OutputPath .\icons.pptx

# 追加到已有演示文稿
pwsh -File scripts\Convert-SvgToShapes.ps1 -SvgPath .\new-icons\ -OutputPath .\existing.pptx -Append
```

## 脚本二 —— `Edit-ExistingPptx.ps1`

四种操作模式，根据所提供的参数自动选择。

### 公共参数（所有模式通用）

| 参数             | 说明 |
|---|---|
| `-TemplatePath`  | 要编辑的已有 `.pptx`（必填）。 |
| `-OutputPath`    | 输出路径。默认为 TemplatePath（原地编辑，自动备份）。 |
| `-NoBackup`      | 跳过创建 `<TemplatePath>.bak.pptx`。 |
| `-Force`         | OutputPath 与 TemplatePath 不同时跳过覆盖确认。 |

### TEXT 模式 —— 仅编辑现有幻灯片的文字（不涉及 SVG）

提供 `-SlideTexts` 而**不**提供 `-SvgPath` 时触发。适用于只需修改文字的封面、标题页和目录页。

`-SlideTexts` 是以 1 开始的幻灯片索引为键的哈希表，每个值描述写入哪个占位符类型的内容：

| 键           | 类型              | 目标占位符 |
|--------------|-------------------|-----------------------------|
| `Title`      | `string`          | 第一个 PP_TITLE / PP_CENTER_TITLE |
| `Subtitle`   | `string`          | 第一个 PP_SUBTITLE |
| `Body`       | `string` 或 `string[]` | Body 占位符文字。若只有一个 body 占位符且提供数组，则各项以换行拼接；若有多个 body 占位符，则按顺序映射。 |
| `Date`       | `string`          | 第一个 PP_DATE |
| `Footer`     | `string`          | 第一个 PP_FOOTER |
| `ShapeTexts` | `object[]`        | 按 `name`/`shapeName` 或 `altText`/`alternativeText` 匹配的文本框；每项需含 `text`。适用于不使用占位符的模板。 |
| `Replacements` | `object[]`      | 对所有含文字的形状（含组合形状）执行查找/替换；每项使用 `find` 和 `replace`。 |

```powershell
# 仅编辑封面（幻灯片 1）和目录（幻灯片 2）—— 不涉及 SVG
pwsh -File scripts\Edit-ExistingPptx.ps1 `
    -TemplatePath .\branded.pptx `
    -OutputPath   .\out.pptx `
    -SlideTexts @{
        1 = @{ Title = '2026 战略报告'; Subtitle = '董事会汇报'; Date = '2026-05' }
        2 = @{ Title = '目录'; Body = @('市场概述','竞争格局','战略规划','执行路径') }
    }
```

对于标题或目录文字使用普通文本框（而非占位符）的精心设计模板，使用 `ShapeTexts` 或 `Replacements`：

```powershell
pwsh -File scripts\Edit-ExistingPptx.ps1 `
    -TemplatePath .\branded.pptx `
    -OutputPath   .\out.pptx `
    -SlideTexts @{
        1 = @{
            ShapeTexts  = @(@{ shapeName = 'CoverHeadline'; text = '2026 战略报告' })
            Replacements = @(@{ find = '{{subtitle}}'; replace = '董事会汇报' })
        }
    }
```

### INSERT 模式 —— 用一个 SVG 替换某张幻灯片的内容

通过 `-TargetSlide` 触发。多个 SVG 从 TargetSlide 开始依次映射到后续幻灯片。

| 参数             | 说明 |
|---|---|
| `-SvgPath`       | `.svg` 文件或目录。 |
| `-TargetSlide`   | 以 1 开始的首个目标幻灯片索引。 |
| `-ContentZone`   | 覆盖 SVG 插入区域：`"左,上,宽,高"`，单位为点。 |
| `-ClearContent`  | 插入前清除中心点落在内容区内的非结构性形状。区域外的模板装饰元素保留。 |
| `-SlideTitles`   | 每个 SVG 对应一个标题（旧版快捷方式）。 |
| `-SlideTexts`    | 每个 SVG 对应的哈希表（键为以 1 开始的 SVG 索引）。 |

```powershell
pwsh -File scripts\Edit-ExistingPptx.ps1 `
    -TemplatePath .\deck.pptx `
    -SvgPath      .\chart.svg `
    -TargetSlide  5 `
    -ClearContent

# 同时设置标题和自定义文本框
pwsh -File scripts\Edit-ExistingPptx.ps1 `
    -TemplatePath .\deck.pptx `
    -SvgPath      .\chart.svg `
    -TargetSlide  5 `
    -ClearContent `
    -SlideTexts @{
        1 = @{
            Title        = '市场竞争格局'
            ShapeTexts   = @(@{ shapeName = 'DataSource'; text = '数据来源：2026 行业报告' })
            Replacements = @(@{ find = '{{year}}'; replace = '2026' })
        }
    }
```

### EXPAND 模式 —— 将内容模板幻灯片复制 N 次

提供 `-SvgPath` 而**不**提供 `-TargetSlide` 时触发。每个 SVG 生成一张新幻灯片，复制所选内容模板幻灯片的所有品牌元素（标题、背景、页脚）；SVG 居中插入内容区（自动检测占位符或默认标题栏下方留边区域）。

| 参数                | 默认值       | 说明 |
|---|---|---|
| `-SvgPath`          | —            | `.svg` 文件或目录。 |
| `-ContentSlide`     | 最后一张幻灯片 | 用作内容模板的幻灯片索引。 |
| `-InsertAfterSlide` | 最后一张幻灯片 | 新幻灯片插入在此索引之后。 |
| `-ContentZone`      | 自动检测     | 覆盖插入区域：`"左,上,宽,高"`，单位为点。 |
| `-ClearContent`     | 关闭         | 插入前清除中心点落在内容区内的非结构性形状。区域外的模板装饰元素保留。 |
| `-SlideTitles`      | —            | 每个 SVG 对应一个标题（旧版快捷方式）。 |
| `-SlideTexts`       | —            | 每个 SVG 对应的哈希表（键为以 1 开始的 SVG 索引）。 |

```powershell
# 以幻灯片 3 为模板复制，插入到幻灯片 2 之后，为每张设置标题
pwsh -File scripts\Edit-ExistingPptx.ps1 `
    -TemplatePath     .\branded.pptx `
    -SvgPath          .\generated\ `
    -ContentSlide     3 `
    -InsertAfterSlide 2 `
    -ClearContent `
    -SlideTexts @{
        1 = @{ Title = '第一章：市场概述' }
        2 = @{ Title = '第二章：竞争格局' }
        3 = @{ Title = '第三章：战略规划' }
    }
```

### MANIFEST 模式 —— 基于 JSON 文件的完整演示文稿工作流

"编辑封面 + 编辑目录 + 添加 N 张内容页"一次调用完成的推荐方式。通过 `-Manifest` 触发。Manifest 中的 SVG 路径相对于 Manifest 文件所在目录解析。

```powershell
pwsh -File scripts\Edit-ExistingPptx.ps1 `
    -TemplatePath .\branded.pptx `
    -Manifest     .\deck.json `
    -OutputPath   .\out.pptx
```

**Manifest 结构**（完整示例见 `examples/deck.json`）：

```json
{
  "edits": [
    {
      "type":     "text",
      "slide":    1,
      "title":    "2026 年度战略报告",
      "subtitle": "董事会汇报",
      "shapeTexts": [
        { "shapeName": "CoverTagline", "text": "内部讨论稿" }
      ],
      "date":     "2026-05"
    },
    {
      "type":  "text",
      "slide": 2,
      "title": "目录",
      "body":  ["市场概述", "竞争格局", "战略规划", "执行路径"]
    },
    {
      "type":          "expand",
      "templateSlide": 3,
      "insertAfter":   2,
      "clearContent":  true,
      "items": [
        { "svg": "slides/s1.svg", "title": "第一章：市场概述" },
        { "svg": "slides/s2.svg", "title": "第二章：竞争格局" },
        { "svg": "slides/s3.svg", "title": "第三章：战略规划" }
      ]
    },
    {
      "type":         "insert",
      "slide":        9,
      "svg":          "slides/closing-chart.svg",
      "title":        "结语",
      "clearContent": true
    }
  ]
}
```

编辑类型及其字段：

| `type`     | 必填字段                         | 可选字段                                                                       |
|------------|----------------------------------|--------------------------------------------------------------------------------|
| `"text"`   | `slide`                          | `title`、`subtitle`、`body`、`date`、`footer`、`shapeTexts[]`、`replacements[]` |
| `"insert"` | `slide`、`svg`                   | `title`、`subtitle`、`body`、`date`、`footer`、`shapeTexts[]`、`replacements[]`、`clearContent`、`contentZone` |
| `"expand"` | `items[]`（每项需含 `svg`）       | `templateSlide`、`insertAfter`、`clearContent`、`contentZone`；每项可含：`title`、`subtitle`、`body`、`date`、`footer`、`shapeTexts[]`、`replacements[]` |

### 内容区自动检测

未提供 `-ContentZone`（或 Manifest 中的 `contentZone`）时：

1. 脚本在参考幻灯片中查找内容占位符（body、object、chart、picture、media），若找到则使用其边界范围。
2. 否则回退到默认值（标题栏下方，约 36 pt 左右边距、108 pt 顶部、18 pt 底部）。

手动获取点值：在 PowerPoint 中右键单击形状 → 大小和位置；英寸数 × 72 即为点数。

`-ClearContent` 使用相同区域：仅删除中心点落在该区域内的非结构性形状。这是针对品牌模板的有意设计：区域外的标题栏、Logo、页脚和装饰元素在扩展时应保留。

### 备份行为

每次运行时自动在原始模板旁创建 `.bak.pptx` 备份（除非指定 `-NoBackup`）。

## Claude 调用这些脚本的方式

1. **使用上方决策树选择正确模式。** 封面/目录 = TEXT；替换单张幻灯片 = INSERT；多张复制内容页 = EXPAND；完整演示文稿 = MANIFEST。
2. 传入前将所有文件路径解析为**绝对路径**。COM 对相对路径处理不稳定。
3. 用户安装了 PowerShell 7 时**优先使用 `pwsh` 而非 `powershell.exe`**，前者默认 UTF-8，可避免命令行中文字符乱码。
4. 对于复杂任务（中文文字 + 多张幻灯片），将 Manifest JSON 写入磁盘后调用 `-Manifest`，完全规避 Shell 引号转义问题。
5. **始终加 `-Force`**。两个脚本在目标文件已存在时会调用 `Read-Host` 等待用户确认；在 Claude 的 Bash 工具等非交互式环境中，这会导致进程永久挂起。`-Force` 跳过该确认。
6. 通过 Bash 运行：
   ```powershell
   pwsh -File "C:\path\scripts\Edit-ExistingPptx.ps1" -TemplatePath "C:\..." -Force ...
   ```
7. 脚本失败时原样输出 stderr。错误信息会区分："PowerPoint 未安装"、"文件未找到"、"幻灯片索引超出范围"、"ExecuteMso 失败"。

## 已知局限

- **PowerPoint 窗口必须可见。** `ExecuteMso` 是功能区命令，隐藏时不可靠。两个脚本均会设置 `Visible = True`。
- **请先关闭其他演示文稿。** COM 会附加到正在运行的实例；已打开的演示文稿可能干扰选择状态。脚本会发出警告但不中止。
- **不支持无头 / CI 环境。** 无头转换请使用 `svg2pptx` 或纯 DrawingML 库。
- **SVG 还原效果取决于 PowerPoint 内置转换器。** 详见上方"SVG 生成规范"。
- **ExecuteMso 时序。** 执行前插入 300 ms 等待。在较慢的机器上可在脚本中调整至 500 ms。
- **每张幻灯片只能插入一个 SVG。** 单张幻灯片多 SVG 布局不在支持范围内。
- **替换占位符文字**使用 `TextRange.Text = …`，可保留布局继承的格式，但可能会抹平 run 级别的覆盖样式（如某个单词被单独设置为红色）。对于此类幻灯片，建议手动编辑以保持精确格式。
- **仅支持 Windows。**

## 验证输出

打开生成的 `.pptx`，右键单击任意插入的形状。若上下文菜单显示**编辑顶点**，或形状可取消组合为独立子形状，则转换成功。若仍显示**转换为形状**，说明该形状未完成转换（通常因为 SVG 格式有误或超出支持子集）；脚本的 stderr 中会有相应记录。
