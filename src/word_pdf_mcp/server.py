from __future__ import annotations

from functools import partial

import anyio
from mcp.server.fastmcp import FastMCP
from mcp.server.fastmcp.exceptions import ToolError

from .errors import ConversionError
from .scheduler import ConversionService

mcp = FastMCP("word-pdf-mcp")
service = ConversionService()


@mcp.tool()
async def word_to_pdf(
    input_path: str,
    output_path: str,
    overwrite: bool = False,
    include_markup: bool = False,
    timeout_seconds: int = 180,
    queue_timeout_seconds: int = 600,
) -> dict:
    """Convert a saved local Word file to PDF using desktop Microsoft Word on Windows.

    Calls serialize across this tool's instances in the same Windows session.
    Source is copied, opened read-only, never saved. Output directory must exist.
    No automatic font repair, OLE refresh, revision acceptance or alternate renderer.
    Default excludes revision/comment markup; include_markup=True exports markup.
    Timeout covers worker lifetime, separately from time spent waiting for the lock.
    A successful PDF still requires visual acceptance for equations and objects.
    """
    try:
        # Keep work alive through client cancellation until owned processes are cleaned up.
        return await anyio.to_thread.run_sync(
            partial(
                service.convert,
                input_path,
                output_path,
                overwrite=overwrite,
                include_markup=include_markup,
                timeout_seconds=timeout_seconds,
                queue_timeout_seconds=queue_timeout_seconds,
            ),
            abandon_on_cancel=False,
        )
    except ConversionError as exc:
        raise ToolError(str(exc)) from exc
    except (OSError, ValueError) as exc:
        raise ToolError(f"file_or_process_error: {exc}") from exc


@mcp.tool()
def word_pdf_status() -> dict:
    """Report platform, mutex name and recovery record location without opening Word."""
    return service.health()


def main() -> None:
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
