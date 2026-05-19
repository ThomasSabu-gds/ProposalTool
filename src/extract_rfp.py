import json
import os
import pdfplumber


# --- Old PDF-only implementation (kept for reference) ---------------------
# def extract_rfp_text(pdf_path: str) -> str:
#     pages = []
#     with pdfplumber.open(pdf_path) as pdf:
#         for page in pdf.pages:
#             text = page.extract_text() or ""
#             if text.strip():
#                 pages.append(text)
#     return "\n\n".join(pages)


def extract_rfp_text(rfp_path: str) -> str:
    """Extract RFP text from a .pdf or .txt file.

    MCP mode writes the raw RFP string to a temporary .txt file and passes
    that path here, so we branch on extension. PDF behavior is unchanged.
    """
    ext = os.path.splitext(rfp_path)[1].lower()
    if ext == ".txt":
        with open(rfp_path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()

    # Default: PDF via pdfplumber (original behavior preserved).
    pages = []
    with pdfplumber.open(rfp_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if text.strip():
                pages.append(text)
    return "\n\n".join(pages)


# ---------------------------------------------------------------------------
# RFP Analysis: extract key facts + condensed knowledge base
# ---------------------------------------------------------------------------

ANALYZE_RFP_SYSTEM = """You are an expert RFP analyst. Analyze the provided RFP text and extract:

Return JSON:
{
  "client_name": "The actual name of the client/authority issuing the RFP (not abbreviations like 'the Authority')",
  "project_name": "Full project name",
  "rfp_reference": "RFP/tender reference number if any",
  "sector": "Industry sector (e.g. Healthcare, Infrastructure, IT, Oil & Gas)",
  "country_region": "Geographic location",
  "key_dates": ["List of important dates mentioned"],
  "scope_summary": "2-3 sentence summary of what the project is about",
  "key_requirements": [
    "Requirement 1",
    "Requirement 2"
  ],
  "evaluation_criteria": ["Criterion 1", "Criterion 2"],
  "methodology_needs": "What methodologies/approaches does the RFP expect from bidders?",
  "team_requirements": "What team/staffing does the RFP expect?",
  "budget_info": "Any budget or financial information mentioned"
}

Rules:
- Extract the ACTUAL client name, not generic references
- If client name is redacted as [Name of Authority], note that
- Be thorough with requirements - these drive proposal content
- Keep each field concise but complete"""


def analyze_rfp(client, deployment: str, rfp_text: str) -> dict:
    """Extract key facts and create a knowledge base from the RFP."""
    # Use first ~40K chars for analysis (covers most RFPs)
    text_for_analysis = rfp_text[:40000]

    print("    Extracting RFP key facts...")
    resp = client.chat.completions.create(
        model=deployment,
        messages=[
            {"role": "system", "content": ANALYZE_RFP_SYSTEM},
            {"role": "user", "content": f"Analyze this RFP:\n\n{text_for_analysis}"},
        ],
        response_format={"type": "json_object"},
        temperature=0.05,
        max_tokens=4096,
    )
    key_facts = json.loads(resp.choices[0].message.content)

    # Build condensed knowledge base from full RFP
    print("    Building RFP knowledge base...")
    knowledge_base = _build_knowledge_base(client, deployment, rfp_text)

    # Build a single structured summary for consistent reference across all slides
    print("    Creating RFP summary for slide agents...")
    rfp_summary = _create_rfp_summary(client, deployment, key_facts, knowledge_base)

    return {
        "key_facts": key_facts,
        "knowledge_base": knowledge_base,
        "rfp_summary": rfp_summary,
        "full_text_length": len(rfp_text),
    }


KNOWLEDGE_BASE_SYSTEM = """Condense this RFP text into a structured knowledge base that can be used to generate proposal content. 

Return JSON:
{
  "sections": [
    {
      "topic": "Section topic name",
      "key_points": ["Point 1", "Point 2"],
      "details": "Important details, numbers, dates, specifications"
    }
  ]
}

Rules:
- Capture ALL substantive requirements, specifications, and criteria
- Preserve exact numbers, dates, percentages, and thresholds
- Group by logical topic (scope, eligibility, evaluation, timeline, etc.)
- Be thorough - this is the ONLY reference for proposal generation
- Maximum 20 sections"""


def _build_knowledge_base(client, deployment: str, rfp_text: str) -> list[dict]:
    MAX_CHARS = 50000
    chunks = _chunk_text(rfp_text, MAX_CHARS)
    all_sections = []

    for i, chunk in enumerate(chunks):
        print(f"    Processing knowledge chunk {i + 1}/{len(chunks)}...")
        try:
            resp = client.chat.completions.create(
                model=deployment,
                messages=[
                    {"role": "system", "content": KNOWLEDGE_BASE_SYSTEM},
                    {"role": "user", "content": f"Condense this RFP chunk ({i+1}/{len(chunks)}):\n\n{chunk}"},
                ],
                response_format={"type": "json_object"},
                temperature=0.05,
                max_tokens=4096,
            )
            result = json.loads(resp.choices[0].message.content)
            all_sections.extend(result.get("sections", []))
        except Exception as e:
            print(f"    Warning: chunk {i + 1} failed: {e}")

    return all_sections


def _chunk_text(text: str, max_chars: int) -> list[str]:
    if len(text) <= max_chars:
        return [text]
    chunks = []
    paragraphs = text.split("\n\n")
    current = ""
    for para in paragraphs:
        if len(current) + len(para) + 2 > max_chars and current:
            chunks.append(current)
            current = para
        else:
            current = current + "\n\n" + para if current else para
    if current:
        chunks.append(current)
    return chunks


RFP_SUMMARY_SYSTEM = """Create a concise executive summary of this RFP that a proposal writer can reference for EVERY slide. Single source of truth about client needs.

Return JSON:
{
  "one_liner": "One sentence describing the RFP",
  "client": "Client name and description",
  "project": "Project name and description",
  "sector": "Industry sector",
  "scope": ["What the client needs - point 1", "point 2"],
  "deliverables": ["Deliverable 1", "Deliverable 2"],
  "evaluation_criteria": ["Criterion 1", "Criterion 2"],
  "timeline": "Key dates and durations",
  "budget": "Budget info if available",
  "key_requirements": ["Requirement 1", "Requirement 2"],
  "methodology_expectations": "What approach the client expects",
  "team_expectations": "What team the client expects"
}"""


def _create_rfp_summary(client, deployment, key_facts, knowledge_base):
    kb_text = ""
    for sec in knowledge_base[:15]:
        kb_text += f"\n{sec.get('topic', '')}: "
        for pt in sec.get("key_points", []):
            kb_text += f"\n  - {pt}"
    try:
        resp = client.chat.completions.create(
            model=deployment,
            messages=[
                {"role": "system", "content": RFP_SUMMARY_SYSTEM},
                {"role": "user", "content": f"Key Facts:\n{json.dumps(key_facts, indent=2)}\n\nKB:\n{kb_text[:4000]}"},
            ],
            response_format={"type": "json_object"},
            temperature=0.05,
            max_tokens=2048,
        )
        return json.loads(resp.choices[0].message.content)
    except Exception as e:
        print(f"    Warning: summary failed: {e}")
        return {"one_liner": "Summary unavailable", **key_facts}
