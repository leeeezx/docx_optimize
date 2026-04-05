# 全局 `.dotm` 改造任务列表

## 一、任务目标

目标链路：

`项目自带 VBA 宏 -> DocxOptimize.dotm -> Startup 安装 -> 脚本稳定调用 -> 面向用户分发`

---

## 二、任务拆解

| 编号 | 任务 | 负责人 | 状态 | 说明 |
|---:|---|---|---|---|
| 1 | 新增安装脚本 `scripts/packaging/install_dotm.ps1` | Codex | 已完成 | 负责复制模板到 Word Startup 并备份旧版本 |
| 2 | 新增卸载脚本 `scripts/packaging/uninstall_dotm.ps1` | Codex | 已完成 | 负责将已安装模板移出 Startup |
| 3 | 新增自检脚本 `scripts/packaging/test_dotm.ps1` | Codex | 已完成 | 负责验证模板存在、Word 可调用宏 |
| 4 | 改造 `02convert_equation_format_MathML_to_OMML.ps1` | Codex | 已完成 | 增加模板存在检查、加载提示、限定宏名调用 |
| 5 | 重写总 README | Codex | 已完成 | 从“导入 Normal”改为“安装全局 `.dotm`” |
| 6 | 新增安装文档 | Codex | 已完成 | 面向最终用户提供安装、卸载、自检说明 |
| 7 | 生成 `DocxOptimize.dotm` 母版 | 你 | 待执行 | 需在 Word 中手工创建或导出成品模板 |
| 8 | 在本机验证 `.dotm` 可加载 | 你 | 待执行 | 需要真实 Word 环境人工确认 |
| 9 | 验证 MathType 宏 `MTCommand_ConvertEqns` | 你 | 待执行 | 这是外部依赖，不在仓库内 |
| 10 | 组装最终发布包 | 你 | 待执行 | 至少包含 `.dotm`、安装脚本和文档 |
| 11 | 可选：补自动构建 `build_dotm.ps1` | 后续 | 待规划 | 用于自动从 `.bas` 构建 `.dotm` |
| 12 | 可选：给 `.dotm` 做数字签名 | 后续 | 待规划 | 面向企业环境时建议做 |

---

## 三、建议执行顺序

执行链：

`先生成 dotm 母版 -> 再运行安装脚本 -> 再做自检 -> 再验证完整流程 -> 最后发布`

| 顺序 | 动作 |
|---:|---|
| 1 | 你在 Word 中生成 `DocxOptimize.dotm` |
| 2 | 将模板放入 `dist/` 或发布包根目录 |
| 3 | 运行 `scripts/packaging/install_dotm.ps1` |
| 4 | 运行 `scripts/packaging/test_dotm.ps1` |
| 5 | 运行 `scripts/math_ops/02convert_equation_format_MathML_to_OMML.ps1` |
| 6 | 如需完整链路，再运行 `03run_equation_pipeline.py` |

---

## 四、当前状态判断

| 模块 | 状态 | 结论 |
|---|---|---|
| 仓库脚本侧 | 已就绪 | 可以支撑全局 `.dotm` 安装式分发 |
| 模板成品侧 | 未完成 | 仍需你生成并放入发布包 |
| 最终用户可用性 | 待验证 | 需你在真实 Word 环境跑通一次 |
