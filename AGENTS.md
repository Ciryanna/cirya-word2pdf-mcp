# Agent handoff rules

This project is a local Windows Word-to-PDF MCP, not an editor or a cloud service.

- Read `docs/CODEX_HANDOFF.md` before modifying process lifecycle or Word automation.
- Never silently use LibreOffice, HTML, screenshot-to-PDF or a print driver as a fallback.
- Do not re-save source documents with python-docx. Preserve the source bytes.
- Hold the named mutex through launch, export, shutdown and verified cleanup. Acquire and release on the same OS thread.
- Never kill all Word processes. Use positively owned PID + creation time. Preserve fail-closed recovery for unknown ownership.
- Keep COM objects in the worker's STA thread; do not pass them to MCP/async threads.
- Do not claim high fidelity based on PDF structure, page count, or mocked tests. Native Word and complex equations need real acceptance samples.
- Run `uv run pytest` and `uv run ruff check .`; Windows tests are necessary for platform-specific changes.
- Keep personal documents, generated PDFs, machine paths and secrets out of git.
- Keep user-facing behavior small: two MCP tools, no HTTP listener, no service installer.
