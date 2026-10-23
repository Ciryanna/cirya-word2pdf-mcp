# Codex 接手说明

## v0.2 npm 入口

当前对外入口为 `bin/cli.mjs`，npm 包名、命令和 MCP 服务名统一为 `cirya-word2pdf-mcp`，版本 `0.2.0`。GitHub 仓库为 `Ciryanna/cirya-word2pdf-mcp`；发布状态以 npm registry 和 GitHub Release 为准。运行链路是 TypeScript MCP → 隐藏 PowerShell STA supervisor → 独立 PowerShell STA Word Worker。Python v0.1 仅保留作回归对照，不打入 npm 包，也没有自动回退到 Python 的行为。

`native/WindowsNative.cs` 通过系统 Add-Type 编译互操作帮助类，不依赖额外安装的 .NET SDK、C++ 工具或 PowerShell 模块。COM 只在 Worker 的 STA 线程操作。`.NET MutexLease` 检查获取/释放线程相同；保留原 mutex 名与 session journal 路径，与 v0.1 协调。

Supervisor 持锁完成恢复、源副本、Worker 等待/超时、验证 Word 退出、调用 Node 只读 PDF 验证、源哈希检查、原子发布和清理。Worker 为已确认归属的 Word 创建 kill-on-close Job。Node→supervisor→Worker 的父进程监控在 C# 后台线程完成，不在线程间传递 COM。

恢复记录继续使用 `pid`、Unix 秒 `created` 和 `name`，兼容 Python v0.1；启动归属未知继续 fail-closed。清理目录先核对绝对路径为本会话 `task-*` 直接子目录，并拒绝重解析点。

新增检查：`npm ci`、`npm run check`、`npm test`、`npm pack --dry-run`。真实 Word 使用 `WORD_PDF_RUN_NATIVE=1` 或 `WORD_PDF_TEST_DOCUMENT`。Python 的 `uv run pytest` 与 `uv run ruff check .` 继续作为回归检查。

特别注意：PowerShell 的 COM 方法列表可能不列出 `ExportAsFixedFormat2`。必须通过 IDispatch.GetIDsOfNames 查询实际接口；只对 MEMBERNOTFOUND/UNKNOWNNAME 判定缺失。原生 API 2 的末尾参数顺序是 `OptimizeForImageQuality, FixedFormatExtClassPtr`，不能按旧 API 直接在末尾追加 Boolean。

发布包必须包含 `lib/`。不要放到历史 Python 构建输出 `dist/`：hatchling 写入的嵌套 `.gitignore` 可能让 npm 漏打 JavaScript。包内容自动测试已覆盖。更新依赖 npm 的 `@latest` + `--prefer-online`，在下次启动生效；没有常驻更新服务或 HTTP 入口。

## 目标与约束

用户在自己的 Windows 电脑使用，含大量高中物理公式和嵌入对象。只做 Word 原生转 PDF。最重要的工程要求是多个 Agent/MCP 实例串行、故障恢复，以及不影响用户日常 Word。

v0.1 最初在 Linux 环境编写，现已完成 Windows + Word 2021 基础原生导出联调，修复了 Job Object 空名称和 Word Application 无 Hwnd 的兼容性问题。基础通过不等于复杂文档验收；继续按 WINDOWS_ACCEPTANCE.md 验收，不要直接宣称所有文档 100% 保真。

## 接手顺序

1. 普通用户运行 `uv sync --extra dev`，执行全部测试。先看 Windows mutex 跨进程测试是否通过。
2. 手动打开 Word，完成激活与首次弹窗，安装文档字体。
3. 调用 MCP `word_pdf_status`；设置 `WORD_PDF_TEST_DOCUMENT`，执行 opt-in 原生 Word 测试。
4. 用用户实际安装的 Codex 客户端接入 stdio，核对工具超时设置。
5. 执行 `WINDOWS_ACCEPTANCE.md` 的真实双客户端、保留日常 Word、超时与崩溃检查。
6. 用真实公式/MathType/OLE 文档与手动参考 PDF 比较。样本不能提交到仓库。

## 必须重点检查

- `DispatchEx` 在用户 Word 版本上的 PID 隔离行为。Word Application 没有 `Hwnd`：当前只对无文档实例临时设置唯一 Caption，通过隐藏 OpusApp 窗口取得 PID，并在 finally 恢复标题，再核对启动前 PID 集合、创建时间与进程名。不能确认唯一身份就停止；不应把这当作对所有 Office 配置的绝对保证。
- Job Object 在客户端进程环境中是否可成功分配（某些宿主已有 Job 限制）。失败应停止并清理专用 Word，不能悄悄降低隔离。
- 创建 Word 与记录 PID/分配 Job 之间有不可原子化窗口。崩溃时保留 dirty journal，后续停止并要求人工检查；不能通过猜测进程修复。
- `ExportAsFixedFormat2` 方法缺失在动态 COM 上是否确实表现为 AttributeError；如果旧版是明确的 DISP_E_MEMBERNOTFOUND/UNKNOWNNAME，可仅为这两类“方法不存在”情况补回退。不要对通用 COM 错误回退。
- Word.Options 更新策略、修订输出，以及插件是否仍触发弹窗或对象刷新。代码在正常退出前恢复修改过的 Options；强杀路径无法保证恢复，Word 独立实例也可能共享用户配置，需实测，必要时使用专用 profile。
- 字体预检尚未实现；OOXML 目前只数 embedded parts，不代表已经检查每个对象或公式。
- 取消调用：`anyio.to_thread.run_sync(abandon_on_cancel=False)` 在正常 MCP 生命周期内继续清理。宿主强杀进程则依赖 Worker 父进程监控、Job 与 journal；实测该路径。
- 当前不保证跨进程 FIFO；只保证互斥，不要在说明中承诺严格提交顺序。
- 慢文档：宿主工具超时需大于 queue timeout + conversion timeout + cleanup。

## 已刻意不做

CLI、HTTP、云部署、任务数据库、常驻 Word、多 Worker 并行、自动 OLE 更新、公式格式转换、完整视觉差异系统。初版每份启动 Word 的额外开销可以接受。

## 下一步可选增强

- 先完成 Windows 原生验收，再处理导出参数微调。
- 在不改变正文的前提下，增加字体/对象 ProgID 清单。
- 增加人工参考 PDF 的像素差异报告与公式区域抽查。
- 有真实长批次数据以后再考虑常驻 Word。
