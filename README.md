# Cirya Word2PDF MCP

本地 Windows Word → PDF MCP。npm 负责分发、启动与版本更新，转换全部通过本机桌面 Microsoft Word COM 完成，文件无需上传。

v0.2 使用 **TypeScript MCP + Windows PowerShell STA + Word COM**。运行时不依赖 Python、uv、C++ 构建工具或额外的 PowerShell 模块。仓库保留 Python v0.1 代码作为回归对照，npm 包不会包含或调用它。

原生链路已在 Word 2021 上实测，但复杂公式、MathType、OLE 和缺字体场景仍需实样视觉验收。成功返回始终包含 `visual_fidelity_verified: false`，页数和 PDF 结构检查不代表 100% 视觉保真。

## 运行要求

- Windows 10/11、Node.js 22+、系统自带的 Windows PowerShell 5.1。
- 安装并激活桌面 Microsoft Word，先手动启动一次并完成首次运行提示。
- 安装文档需要的中文、数学与 MathType 字体及组件。
- 在用户的交互桌面会话运行，不作为 Windows 服务部署。

只有两个 MCP 工具：`word_pdf_status`、`word_to_pdf`。使用 stdio，不开 HTTP 端口，不需要 API Key。

## 从当前源码运行

```powershell
npm ci
npm run build
npm start
```

MCP 启动后等待标准输入是正常行为。可用 `node bin/cli.mjs --help` 查看基本要求。

开发期间的 MCP 配置使用本地入口，替换项目路径：

```json
{
  "mcpServers": {
    "word-pdf": {
      "command": "node",
      "args": ["<项目绝对路径>/bin/cli.mjs"]
    }
  }
}
```

## npm 安装与更新

包名、命令和 MCP 服务名均为 `cirya-word2pdf-mcp`，客户端配置 ID 推荐使用 `word_pdf`。

普通用户推荐全局安装，同一台 Windows 上的不同项目与 MCP 客户端可以共用：

```powershell
npm install -g cirya-word2pdf-mcp@latest
```

再次执行这条命令即可升级，升级后重启 MCP。全局安装不会自动更新。

Pi 使用 `~/.pi/agent/mcp.json`，将以下条目合并到现有 `mcpServers` 后执行 `/reload`；`timeout` 单位为秒：

```json
{
  "mcpServers": {
    "word_pdf": {
      "command": "cmd",
      "args": ["/d", "/c", "cirya-word2pdf-mcp"],
      "timeout": 900,
      "exposure": "direct"
    }
  }
}
```

如果希望启动时查询新版，也可以直接由 npx 获取并启动：

```json
{
  "mcpServers": {
    "word_pdf": {
      "command": "cmd",
      "args": ["/d", "/c", "npx", "--yes", "--prefer-online", "cirya-word2pdf-mcp@latest"]
    }
  }
}
```

`@latest` 加 `--prefer-online` 在每次启动时查询发布版本，由 npm 下载并缓存，随后在本机运行。新版本在下一次重启 MCP 时生效；正在执行的转换不会被热替换。获取更新需要能访问 npm registry。

需要固定版本或回退时，将 `@latest` 换成明确版本，例如 `@0.2.0`。不需要开发额外的常驻更新服务。

包管理与更新行为参考 [npm exec 官方文档](https://docs.npmjs.com/cli/commands/npm-exec/)。

## 调用工具

先调用 `word_pdf_status` 查看会话状态。它不会为了检查环境而启动 Word；有 pending 记录可能表示正在转换，也可能需要恢复，不应直接删除。

```json
{
  "input_path": "<本地绝对路径>/试卷.docx",
  "output_path": "<本地绝对路径>/试卷.pdf",
  "overwrite": false,
  "include_markup": false,
  "timeout_seconds": 180,
  "queue_timeout_seconds": 600
}
```

输入支持 `.doc`、`.docx`、`.docm`。输出目录须已存在；默认不覆盖。默认导出正文、不带修订批注标记，不接受或拒绝修订。源文档按字节复制，只读打开副本，不保存源文件。

转换超时包括独立 Worker 的启动、导出和退出；锁等待超时另计。宿主工具调用超时应大于排队 + 转换 + 清理时间，默认设置建议至少 900 秒。客户端取消请求后，转换和清理会继续执行到可安全结束。

不提供密码输入和 Protected View 交互；打开失败或弹窗阻塞会返回错误或超时。所有路径都属于运行 MCP 的本机。

## 原生转换与进程保护

1. TypeScript 提供 MCP 协议与参数校验，为每次请求启动隐藏的 PowerShell supervisor。
2. Supervisor 在同一 STA 线程持有 Windows 命名互斥锁，覆盖恢复、启动、导出、退出、验证、发布与清理。
3. 用 SHA-256 验证按字节复制的源文件，再由独立 STA Worker 创建 Word COM 实例。
4. 使用唯一临时 Caption 定位隐藏 OpusApp 窗口，立即恢复标题；核对 PID、创建时间、进程名和空文档集合，拒绝复用已有 Word。
5. 只将确认归属的 Word 放入 `KILL_ON_JOB_CLOSE` Job Object。Worker 与 supervisor 均监控父进程身份。
6. 禁用宏、自动外链和打印时字段更新，只读打开副本，重新分页。
7. 通过 IDispatch 查询实际 API：优先 `ExportAsFixedFormat2`，开启原始图像质量；只有接口确实不存在时使用原生 `ExportAsFixedFormat` 并告警。真实导出错误不重试其他 API 或引擎。
8. 关闭专用 Word、核实退出，再用 pdf-lib 只读检查 PDF 结构和 Word/PDF 页数，核对源哈希，原子发布 PDF。
9. 验证清理完成后才释放锁。未知归属保留 journal 并返回 `recovery_required`，绝不全局结束 Word。

命名锁沿用 `Local\Elwina.WordPdf.Worker.v1`，与 Python v0.1 在同一 Windows 会话内协调。没有跨登录会话互斥和严格 FIFO 保证。

不会使用 LibreOffice、HTML、截图、打印驱动、python-docx 重存或其他渲染降级；不会刷新 OLE、转换公式格式、升级兼容模式或自动修复字体。`BitmapMissingFonts` 无法补回本机缺少的字体。PDF 不保留 OLE 的交互编辑能力。

## 开发与验证

```powershell
npm ci
npm run check
npm test
```

普通测试覆盖真实 MCP stdio、Windows 互斥/遗弃锁、Job 回收、PID 复用保护、未知归属阻断、PDF 验证、原子输出及 npm 包内容。真实 Word 测试须明确开启：

```powershell
$env:WORD_PDF_RUN_NATIVE = '1'
npm test
```

这会创建临时的中文和 OMML 公式 DOCX，执行真实 Word 转换并清理样本。也可指定自有文档；源文件不会修改：

```powershell
$env:WORD_PDF_TEST_DOCUMENT = '<本地绝对路径>/真实样本.docx'
npm run test:word
```

历史 Python 回归检查：

```powershell
uv run pytest
uv run ruff check .
```

打包前检查：

```powershell
npm run check
npm test
npm pack --dry-run
npm pack
```

npm 包仅包含入口、编译后的 JavaScript、PowerShell/.NET 互操作代码和说明，不包含测试、个人文档、Python 或本机虚拟环境。打包不会发布到 registry。

参阅 [Windows 验收](docs/WINDOWS_ACCEPTANCE.md)、[验证记录](docs/VALIDATION.md)、[恢复说明](docs/RECOVERY.md)、[开发交接](docs/CODEX_HANDOFF.md)。

## 源码布局

| 路径 | 职责 |
| --- | --- |
| `bin/cli.mjs` | npm 可执行入口 |
| `node/server.ts` | 两个 stdio MCP 工具 |
| `node/service.ts` | 隐藏启动 supervisor、请求边界与最终时限 |
| `node/verify-pdf.ts` | PDF 只读解析与页数检查 |
| `native/supervisor.ps1` | 互斥、恢复、字节副本、超时、发布与清理 |
| `native/worker.ps1` | STA Word COM 原生导出 |
| `native/WindowsNative.cs` | Windows Job、进程身份、父监控、互斥和 COM 方法探测 |
| `native/common.ps1` | 日志、哈希、恢复、嵌入部件清点 |
| `test/` | npm/Windows 自动测试与真实 Word opt-in 验证 |
| `src/word_pdf_mcp/`, `tests/` | Python v0.1 回归对照，不进入 npm 运行时 |
