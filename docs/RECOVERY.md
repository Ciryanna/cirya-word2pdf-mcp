# 异常恢复

v0.2 npm 版的进程链是 Node MCP → PowerShell supervisor → PowerShell Worker → 专用 Word。supervisor 和 Worker 都监控自己的父进程身份；Word Job 属于 Worker。状态目录、mutex 名及身份记录格式与 Python v0.1 共用，便于升级期间互斥和恢复。

正常超时会终止本任务 Worker；Job Object 回收其拥有的 Word，外层持锁完成核实后才继续后续任务。

`recovery_required` 表示清理无法被可靠确认。系统保留 `pending.json`，阻止新任务继续启动 Word。这样可以避免将日常 Word 误认为遗留任务。

1. 停止所有使用本工具的 MCP 客户端/进程，确保没有转换正在执行。
2. 从报错或 `word_pdf_status` 获取状态目录。默认在 `%LOCALAPPDATA%\WordPdfMcp\session-<id>`。
3. 阅读 `pending.json` 和对应任务目录中的 `state.json`、`worker.log`。核对 PID、创建时间、进程名；不要只凭 PID 结束进程。
4. 如果阶段停在 `starting_word`，没有可核实的 Word 身份，人工检查 Word 窗口/任务管理器，先保存自己的工作，再关闭确认遗留的自动化 Word。无法分辨时，保存工作后退出所有 Word 是人工恢复选项，不是程序自动行为。
5. 确认对应 Worker 和遗留自动化 Word 均已退出后，手动删除该状态目录里的 `pending.json`，再重新启动 MCP。
6. 可以删除对应 `task-*` 目录以及记录的 `.wordpdf-task-*.pdf` 临时输出。不要删除原文档或目标 PDF。遗留任务目录含原文件副本，按个人文档处理。

进程归属已明确的遗留任务，下一次持锁调用会尝试自动清理。未知归属必须人工确认；没有“一键清除所有 Word”的接口。
