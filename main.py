"""啟動 AutoCAD MCP 伺服器的入口點"""
import asyncio
from autocad_mcp.server import main

if __name__ == "__main__":
    asyncio.run(main())
