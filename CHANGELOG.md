# Changelog

## 0.2.0

- Provide two local stdio MCP tools through a TypeScript/npm entry point: `word_to_pdf` and `word_pdf_status`.
- Run native desktop Microsoft Word COM in a dedicated Windows PowerShell STA worker, without a Python runtime dependency.
- Prefer `ExportAsFixedFormat2` with original image quality; use the older native Word export API only when the newer method is absent.
- Preserve source bytes, verify SHA-256 and PDF page count, and publish output atomically without overwriting by default.
- Serialize conversions across processes within the Windows session and retain PID/creation-time ownership checks, kill-on-close Job protection, and fail-closed recovery.
- Verify real Word 2021 export with Chinese and OMML samples, independent competing conversion processes, and an npm-packed stdio client.

Complex equations, MathType/OLE and missing-font scenarios still require real visual acceptance samples. Successful export does not claim verified visual fidelity.
