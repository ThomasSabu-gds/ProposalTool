"""Quick smoke test for the MCP engine. Not part of the MCP server itself."""
import json
import os
import sys

_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _THIS_DIR)

import mcp_engine

prompt_path = os.path.join(_THIS_DIR, "prompt.txt")
with open(prompt_path, "r", encoding="utf-8") as f:
    text = f.read()

result = mcp_engine.generate_proposal(text, user_prompt="")
summary = {
    "proposal_id": result["proposal_id"],
    "output_path": result["output_path"],
    "slides_built": result["slides_built"],
    "timing_seconds": result["timing_seconds"],
    "client": result["plan"].get("client_name"),
    "sector": result["plan"].get("sector"),
    "section_count": len(result["plan"].get("sections", [])),
    "iter_summary": [
        {"slide": r["slide_index"] + 1,
         "type": r["section_type"],
         "iter": r["iterations"],
         "ok": r["accepted"]}
        for r in result["slide_reports"]
    ],
}
print(json.dumps(summary, indent=2))
