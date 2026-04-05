# READMe

## 一、项目定位

本项目是对 `anthropics/skills.git/skills/docx` 的增强版本，核心目标是补齐 MathType 公式处理链路：

`MathType OLE -> MathML -> OMML -> 后续提取/转换`

当前版本开始将项目自带 VBA 宏的分发方式从“手工导入 `Normal.dotm`”迁移为“安装全局 `.dotm` 宏模板/加载项”。

---

## 二、能力边界

| 能力 | 依赖来源 | 说明 |
|---|---|---|
| `PlainMathMLToEquation` | 本项目提供的 `DocxOptimize.dotm` | 负责 `MathML -> OMML` |
| `MTCommand_ConvertEqns` | 用户本机 MathType | 负责 `OLE -> MathML` |
| Word 文档自动处理 | 本项目 PowerShell / Python 脚本 | 负责串联整条流程 |

关键结论：

1. `DocxOptimize.dotm` 只能封装项目自带宏。
2. MathType 宏 `MTCommand_ConvertEqns` 仍然要求用户本机正确安装 MathType。

---

## 三、运行环境

| 项目 | 是否必须 | 说明 |
|---|---:|---|
| Windows | 是 | 依赖 Word COM 自动化 |
| Microsoft Word | 是 | 宏在 Word 中执行 |
| Python 3 | 是 | 运行总控脚本时需要 |
| PowerShell | 是 | 执行安装脚本与转换脚本 |
| MathType | 视场景而定 | 仅完整流程第 1 步需要 |
| 启用 VBA 宏 | 是 | 否则 `.dotm` 不会运行 |

MathType 设置参考图：

| 图示 | 路径 |
|---|---|
| 图 1 | [fig/mathtype_option01.png](/E:/Dev/skill_creat/docx_optimize/fig/mathtype_option01.png) |
| 图 2 | [fig/mathtype_option02.png](/E:/Dev/skill_creat/docx_optimize/fig/mathtype_option02.png) |

---

## 四、推荐安装方式

推荐链路：

`准备 DocxOptimize.dotm -> 运行安装脚本 -> 启动 Word -> 启用宏 -> 运行自检`

安装命令：

```powershell
pwsh -ExecutionPolicy Bypass -File scripts\packaging\install_dotm.ps1
```

自检命令：

```powershell
pwsh -ExecutionPolicy Bypass -File scripts\packaging\test_dotm.ps1
```

详细说明见：

| 文档 | 作用 |
|---|---|
| [docs/安装说明.md](/E:/Dev/skill_creat/docx_optimize/docs/安装说明.md) | 面向用户的安装与使用步骤 |
| [docs/全局_dotm_任务列表.md](/E:/Dev/skill_creat/docx_optimize/docs/全局_dotm_任务列表.md) | 当前全局 `.dotm` 改造任务拆解 |

---

## 五、常用命令

### 1. 完整公式转换流程

执行链：

`检查 OLE -> 调用 MathType 宏 -> 调用 DocxOptimize.dotm 宏 -> 输出结果`

```powershell
python scripts\math_ops\03run_equation_pipeline.py input.docx --out output.docx
```

### 2. 只执行 `MathML -> OMML`

```powershell
pwsh -ExecutionPolicy Bypass -File scripts\math_ops\02convert_equation_format_MathML_to_OMML.ps1 -DocxPath input.docx -OutPath output.docx
```

### 3. 卸载全局模板

```powershell
pwsh -ExecutionPolicy Bypass -File scripts\packaging\uninstall_dotm.ps1
```

---

## 六、自查链路

自查流程：

`确认模板已安装 -> 确认 Word 能看到宏 -> 再跑转换脚本`

| 检查项 | 动作 |
|---|---|
| 模板文件是否存在 | 查看 `%AppData%\Microsoft\Word\STARTUP\DocxOptimize.dotm` |
| 宏是否可见 | 打开 Word 后按 `Alt + F8`，检查 `PlainMathMLToEquation` |
| 模板工程是否加载 | 按 `Alt + F11`，查看 `Project (DocxOptimize.dotm)` |
| MathType 宏是否可见 | 在 Word 宏列表中确认 `MTCommand_ConvertEqns` |

若第 2 步脚本提示未检测到 `DocxOptimize.dotm`，先重新执行安装脚本，再重启 Word。
