import asyncio
import sys

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client


def test_real_stdio_handshake_tools_and_platform_error():
    async def check():
        parameters = StdioServerParameters(
            command=sys.executable, args=["-m", "word_pdf_mcp.server"]
        )
        async with stdio_client(parameters) as (read, write):
            async with ClientSession(read, write) as session:
                initialized = await session.initialize()
                assert initialized.serverInfo.name == "word-pdf-mcp"
                tools = await session.list_tools()
                assert {tool.name for tool in tools.tools} == {
                    "word_to_pdf",
                    "word_pdf_status",
                }
                status = await session.call_tool("word_pdf_status", {})
                assert not status.isError
                if sys.platform != "win32":
                    failure = await session.call_tool(
                        "word_to_pdf", {"input_path": "missing.docx", "output_path": "missing.pdf"}
                    )
                    assert failure.isError
                    assert "windows_required" in failure.content[0].text

    asyncio.run(check())
