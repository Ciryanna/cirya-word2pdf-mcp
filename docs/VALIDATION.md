# 验证记录

## v0.2 npm 本机验证

环境：Windows、Node.js 24.18.0、Windows PowerShell 5.1、Word 2021（16.0.17932）。以下已实际执行：

| 检查 | 结果 |
| --- | --- |
| `WORD_PDF_RUN_NATIVE=1` + `npm test` | 8 项通过，无跳过，含真实 Word |
| `npm run check` | TypeScript 检查通过 |
| npm stdio | 真实 initialize、两工具枚举、状态、输入错误返回通过 |
| PowerShell/.NET 内核保护 | 跨进程互斥、遗弃锁、Job 只清理所属进程、PID 复用防误杀、未知归属阻断通过 |
| 输出保护 | 默认不覆盖与显式原子覆盖通过；PDF 头、解析和页数不匹配拒绝通过 |
| 真实 Word 中文/OMML 样本 | 两页，`ExportAsFixedFormat2`，Producer 为 Microsoft Word 2021 |
| 两个独立 supervisor 争用同一输出 | 一个成功，另一个 output_exists；源哈希不变，无新增 Word 进程和 pending 遗留 |
| npm 包内容 | 包含 JavaScript、PowerShell、C# 互操作代码，排除 Python、测试和个人文档 |
| `npm pack` | 本地 tgz 构建通过，未发布 registry |
| 从 tgz 经 npm exec/npx 运行 | 从临时工作目录启动，握手、两工具、真实 Word API 2 导出、源哈希和清理通过 |
| 历史 Python 回归 | `uv run pytest` 24 passed / 1 skipped，`uv run ruff check .` 通过 |
| 远端 CI | Windows Node 22/24 工作流已准备，未执行远端 CI |

npm 运行链路不依赖或调用 Python。pdf-lib 只读取原生 Word 生成的 PDF，验证时关闭元数据更新，不重存 PDF。

尚未完成：人工参考 PDF 逐页视觉对照、复杂公式/MathType/OLE、缺字体、真实用户未保存 Word 保留、Word 转换阶段的强杀/超时整链验收。基础测试与合成公式样本不能证明所有文档视觉保真，成功结果保留 `visual_fidelity_verified: false`。

## v0.1 初版交接记录

日期：2026-10-23。开发环境：Linux / Python 3.12.14。

| 检查 | 结果 |
| --- | --- |
| `pytest -q` | 18 passed, 4 skipped |
| `ruff check .` | 通过 |
| `uv build` | sdist 和 wheel 构建成功 |
| 真实 stdio MCP | 子进程启动、initialize、list_tools、status、错误返回通过 |
| Windows 命名锁跨进程互斥/遗弃恢复 | 测试已实现；本环境跳过 |
| Windows Job Object 回收 | 测试已实现；本环境跳过 |
| 真实 Word COM 转换 | opt-in 测试已实现；本环境无 Word，跳过 |
| 公式/对象视觉保真 | 尚未测试，必须在用户电脑上验收 |
| GitHub Actions | 工作流已准备；不能据此声称远端 CI 已通过 |

便携测试覆盖：并发调用串行、获取锁后到清理结束才释放、转换超时、启动失败、未知归属阻止下一任务、PID 复用保护、现有 PDF 不覆盖、页数不一致不提交、非破坏性对象部件清点、导出参数、旧 API 回退与真实导出错误不回退。

本次依赖：MCP 1.30.0、pypdf 6.21.2、psutil 7.2.2、pytest 9.1.1、ruff 0.17.2；完整解析结果在 `uv.lock`。pywin32 为 Windows 专属依赖，未在 Linux 执行。
