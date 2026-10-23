# Windows 原生验收清单

这些检查需要真实 Windows + Microsoft Word。自动测试的 mock 不能替代。

## 已完成的基础联调

v0.2 npm 版：Windows、Node.js 24.18.0、Windows PowerShell 5.1、同机 Word 2021 已完成真实中文/OMML 两页样本原生导出，实际 API 为 `ExportAsFixedFormat2`，Producer 为 Microsoft Word 2021。验证源哈希、默认不覆盖、无新增 Word 进程及无 pending 遗留；npm stdio、原子输出、命名锁、Job、PID 复用、未知归属与包内容测试通过。详情见 VALIDATION.md；不据此宣称复杂公式视觉保真。

以下为此前 Python v0.1 基础联调记录：

Windows、Python 3.14.6、桌面 Word 2021（Version 16.0，Build 16.0.17932）已验证：

- 全量自动测试包含一次真实 Word 转换，原生导出成功。
- 独立 MCP stdio 客户端完成握手、状态查询、`word_to_pdf` 调用。
- 临时 DOCX 样本包含中文、原生 OMML 分式与上标、显式分页；Word 与 PDF 均为两页，PDF 正文可读取。
- 实际使用 `ExportAsFixedFormat2`，PDF Producer 为 Microsoft Word 2021。
- 源文件 SHA-256 不变；转换后 Word 进程集合恢复到测试前状态。
- 跨进程互斥、遗弃锁和 Job Object 退出清理测试通过；静态检查与打包通过。

修复了 `CreateJobObject(None, None)` 在当前 pywin32 上报错的问题，改用空字符串创建无名 Job；同时用唯一临时 Caption + 隐藏 OpusApp 窗口核对 PID，替代不存在的 `Word.Application.Hwnd`。临时标题在 finally 中恢复，含已有文档的实例会拒绝使用。

这些结果只证明基础原生转换链路可用。下面的真实双客户端、用户 Word 保留、强杀恢复、MathType/OLE、缺字体与人工参考 PDF 视觉比较仍需逐项验收；合成样本的页数和文本检查不构成视觉保真证明。

## 完整验收场景

| 场景 | 通过标准 |
| --- | --- |
| 一个普通中文 DOCX | 导出成功，源文件哈希不变，PDF 可读，Word/PDF 页数一致 |
| 保留一个手动打开且含未保存修改的 Word | 转换前后该窗口、内容和未保存状态不受影响 |
| 两个独立 MCP 客户端同时提交 | 不发生本工具 Word 转换重叠；第二份等待前一份清理完 |
| 两个任务输出同一路径，默认不覆盖 | 一个成功，另一个返回 output_exists；已有 PDF 不被覆盖 |
| 锁等待超时 | 返回 queue_timeout，没有为该请求启动 Word |
| 转换时限耗尽 | 只回收自己的 Worker/Word；后续任务可继续或明确 recovery_required |
| 强杀 Worker | 已分配 Job 的专用 Word 被回收，后续任务不与遗留进程重叠 |
| 强杀 MCP 主进程 | Worker 父监控退出；Job 回收 Word；下次恢复 journal |
| Word 创建阶段崩溃 | 无法确认身份时停止后续任务，不猜测、不误杀 |
| API 不支持 ExportAsFixedFormat2 | 明确回退原生 ExportAsFixedFormat 并返回警告 |
| 原生导出报错 | 返回错误，不改用其他引擎，不覆盖已有 PDF |
| 修订/批注 | 默认正文与 include_markup=True 两种模式符合手动参考 |
| 原生公式 | 分式、上下标、根号、矩阵、希腊字母、跨行公式无裁切/漏字 |
| MathType/旧式公式 | 字体、比例、符号、基线与人工参考一致 |
| 嵌入/浮动对象 | EMF/WMF、图表、组合对象、文本框位置与分页符合参考 |
| 缺少字体或对象组件 | 不声称保真；记录问题，补环境后重新验收 |

参考 PDF 必须先由人工确认的正确环境生成。逐页渲染比较，重点放大公式区域。不要仅凭页数或文本提取判断保真。
