"""Model Context Protocol server for Proposal Assist.

Exposes two tools over stdio JSON-RPC:

  - generate_proposal_from_text(rfp_text, user_prompt?)
        Drives the slide-master <-> verifier loop and writes a PPTX to runs/.

  - generate_proposal_from_file(prompt_path, user_prompt?)
        Convenience wrapper: reads the file (defaults to ./prompt.txt) and
        calls the same engine.

Designed to be wired into Claude Desktop / Cursor via:

  {
    "mcpServers": {
      "proposal-assist": {
        "command": "<repo>/venv/Scripts/python.exe",
        "args":    ["<repo>/mcp_server.py"]
      }
    }
  }

The server does NOT modify the existing pipeline; it imports helpers from
src/ and writes output to runs/.
"""

from __future__ import annotations

import os
import sys
import traceback
from typing import Any

# Make src/ importable so the engine can pull helpers without modifying anything.
_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
_SRC_DIR = os.path.join(_THIS_DIR, "src")
if _SRC_DIR not in sys.path:
    sys.path.insert(0, _SRC_DIR)

from mcp.server.fastmcp import FastMCP  # noqa: E402

import mcp_engine  # noqa: E402

mcp = FastMCP("proposal-assist")


def _serialize(result: dict[str, Any]) -> dict[str, Any]:
    """Drop heavy fields from the engine result so the MCP response stays small."""
    return {
        "proposal_id": result["proposal_id"],
        "template_used": result["template_used"],
        "output_path": result["output_path"],
        "slides_built": result["slides_built"],
        "timing_seconds": result["timing_seconds"],
        "message": result["message"],
        "plan_summary": {
            "client_name": result["plan"].get("client_name"),
            "project_name": result["plan"].get("project_name"),
            "sector": result["plan"].get("sector"),
            "section_count": len(result["plan"].get("sections", [])),
        },
        "slide_reports": [
            {
                "slide_index": r["slide_index"],
                "section_type": r["section_type"],
                "iterations": r["iterations"],
                "accepted": r["accepted"],
                "remaining_issue_count": len(r["remaining_issues"]),
            }
            for r in result["slide_reports"]
        ],
    }


@mcp.tool()
def generate_proposal_from_text(rfp_text: str, user_prompt: str = "") -> dict[str, Any]:
    """Generate a populated proposal PPTX from raw structured RFP text.

    Args:
        rfp_text: The RFP analysis / proposal brief (free-form text). The engine
            parses this into a slide plan and populates it into the default
            Healthcare template via a SlideMaster <-> Verifier loop.
        user_prompt: Optional extra user instructions appended to the parser
            prompt (e.g. "Focus on PPP advisory experience").

    Returns:
        Summary dict including proposal_id, output_path under ./runs/, and
        per-slide verifier reports.
    """
    if not rfp_text or not rfp_text.strip():
        raise ValueError("rfp_text is required and must be non-empty.")
    try:
        result = mcp_engine.generate_proposal(rfp_text, user_prompt=user_prompt)
        return _serialize(result)
    except Exception as e:
        traceback.print_exc()
        raise RuntimeError(f"Pipeline failed: {e}") from e


@mcp.tool()
def generate_proposal_from_file(
    prompt_path: str = "prompt.txt",
    user_prompt: str = "",
) -> dict[str, Any]:
    """Same as generate_proposal_from_text, but reads the RFP text from a file.

    Args:
        prompt_path: Path to the file. Relative paths are resolved against the
            repository root. Defaults to ``prompt.txt``.
        user_prompt: Optional extra user instructions appended to the parser
            prompt.
    """
    full = prompt_path if os.path.isabs(prompt_path) else os.path.join(_THIS_DIR, prompt_path)
    if not os.path.isfile(full):
        raise FileNotFoundError(f"Prompt file not found: {full}")
    with open(full, "r", encoding="utf-8", errors="ignore") as f:
        text = f.read()
    return generate_proposal_from_text(text, user_prompt=user_prompt)


@mcp.tool()
def list_runs() -> dict[str, Any]:
    """List previously generated proposal PPTX files in ./runs/."""
    runs_dir = os.path.join(_THIS_DIR, "runs")
    if not os.path.isdir(runs_dir):
        return {"runs": []}
    items = []
    for name in sorted(os.listdir(runs_dir)):
        full = os.path.join(runs_dir, name)
        if os.path.isfile(full) and name.lower().endswith(".pptx"):
            stat = os.stat(full)
            items.append({
                "name": name,
                "path": full,
                "size_bytes": stat.st_size,
                "modified_ts": stat.st_mtime,
            })
    return {"runs": items}



if __name__ == "__main__":
    mcp.run(transport="streamable-http")


