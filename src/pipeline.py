"""Agentic pipeline with Knowledge Base retrieval and RFP summary."""

import os
import re
import json
import time

from pptx import Presentation

from extract_rfp import extract_rfp_text, analyze_rfp
from knowledge_base import KnowledgeBase, classify_section
from generate_pptx import generate_pptx_from_enrichments
from verifier import verify_and_fix

# Singleton KB — built once with the same client/embedding config, shared
# across requests. `sync_index()` is incremental, so subsequent calls are cheap.
_kb_instance: KnowledgeBase = None
_kb_lock = __import__("threading").RLock()


def get_kb(
    datalake_dir: str,
    llm_client=None,
    llm_deployment: str = None,
    embedding_client=None,
    embedding_deployment: str = None,
) -> KnowledgeBase:
    """Return the singleton KB. Builds + syncs lazily on first call."""
    global _kb_instance
    with _kb_lock:
        if _kb_instance is None:
            print("  [KB] Initializing persistent knowledge base...")
            _kb_instance = KnowledgeBase(
                datalake_dir,
                embedding_client=embedding_client,
                embedding_deployment=embedding_deployment,
                llm_client=llm_client,
                llm_deployment=llm_deployment,
            )
            _kb_instance.sync_index()
        return _kb_instance


# Back-compat alias for any callers still importing the old name
_get_kb = get_kb


def run_pipeline(
    client, deployment, rfp_path, template_path, output_dir,
    user_prompt="", on_progress=None,
    embedding_client=None, embedding_deployment=None,
) -> dict:
    def _emit(step, name, msg):
        if on_progress:
            on_progress(step, name, msg)
        print(f"  [{step}/6] {name}: {msg}")

    t0 = time.time()

    # ── Step 1: Knowledge base (persistent, incremental sync if needed) ──
    _emit(1, "Analyze", "Loading knowledge base...")
    datalake_dir = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "datalake"
    )
    kb = get_kb(
        datalake_dir,
        llm_client=client, llm_deployment=deployment,
        embedding_client=embedding_client, embedding_deployment=embedding_deployment,
    )

    # ── Step 2: Deep RFP Analysis + Summary ──
    _emit(1, "Analyze", "Extracting RFP content...")
    rfp_text = extract_rfp_text(rfp_path)
    rfp_analysis = analyze_rfp(client, deployment, rfp_text)
    kf = rfp_analysis["key_facts"]
    rfp_summary = rfp_analysis["rfp_summary"]

    client_name = kf.get("client_name", "")
    project_name = kf.get("project_name", "")
    rfp_sector = kf.get("sector", rfp_summary.get("sector", "Unknown"))
    if not client_name or client_name.startswith("[") or "name of" in client_name.lower():
        client_name = rfp_summary.get("client", "The Authority")
        if not client_name or client_name.startswith("["):
            client_name = "The Authority"
    if not project_name or "***" in project_name:
        project_name = rfp_summary.get("project", "The Project")

    _emit(1, "Analyze", f"Client: {client_name} | Sector: {rfp_sector}")
    _emit(1, "Analyze", f"Summary: {rfp_summary.get('one_liner', 'N/A')}")

    # Format the RFP summary as text (this goes to EVERY slide agent)
    rfp_summary_text = _format_rfp_summary(rfp_summary)

    # ── Step 3: Read template slides ──
    _emit(2, "Compose", "Reading template slides...")
    template_slides = _extract_template_slides(template_path)
    _emit(2, "Compose", f"{len(template_slides)} slides")

    # ── Step 4: Per-slide agent with KB retrieval ──
    _emit(3, "Format", "Running per-slide content agent with KB retrieval...")
    content_map = {}
    enriched_count = 0

    # Load the template's persisted TOC so the agent knows each slide's role
    # (cover / experience / methodology / divider / ...). The Oil & Gas template
    # has many visually-rich but textually-sparse slides; without TOC purpose
    # they were being skipped, producing an empty deck.
    template_toc = kb.get_toc(os.path.basename(template_path)) or {}
    toc_purpose_by_idx = {
        int(s.get("index", -1)): s.get("purpose", "other")
        for s in template_toc.get("slides", [])
    }
    today = time.strftime("%d %B %Y")

    for slide in template_slides:
        idx = slide["index"]
        # Previously: `if total_words < 5: continue` — this filter caused the
        # empty-deck symptom on Oil & Gas. We now let the agent see every
        # slide and decide whether to fill / generate / keep based on its
        # TOC-assigned purpose.

        # Classify this slide's section type — prefer the persisted TOC purpose
        toc_purpose = toc_purpose_by_idx.get(idx, "")
        section_type = (
            _normalize_purpose(toc_purpose)
            if toc_purpose
            else classify_section(slide["title"], slide["full_text"])
        )

        # Skip enrichment for section dividers / starter slides. These are
        # visual chapter markers (e.g. "01 / Why Us", "02 / Our Experiences").
        # Enriching them dumps paragraphs into the empty Rectangle shape, which
        # destroys the design. The following content slide already gets that
        # content. We also detect them heuristically so non-tagged templates
        # don't slip through.
        if toc_purpose in ("section_header", "divider") or _is_section_header_slide(slide):
            _emit(3, "Format", f"  Slide {idx+1} [{section_type}]: kept (section divider)")
            continue

        # Retrieve relevant knowledge from the DB
        kb_content = kb.get_closest_sector_content(
            rfp_sector=rfp_sector,
            section_type=section_type,
            query_text=slide["full_text"][:300],
            n_results=3,
        )

        result = _slide_agent(
            client, deployment, slide,
            rfp_summary_text=rfp_summary_text,
            kb_content=kb_content,
            section_type=section_type,
            slide_purpose=toc_purpose,
            client_name=client_name,
            project_name=project_name,
            rfp_sector=rfp_sector,
            user_prompt=user_prompt,
            today_date=today,
            tender_number=kf.get("tender_number", "") or "",
            country_region=kf.get("country_region", "") or "",
        )

        if result and result.get("action") == "modify":
            content_map[idx] = result["shapes_content"]
            enriched_count += 1
            _emit(3, "Format", f"  Slide {idx+1} [{section_type}]: ENRICHED")
        else:
            _emit(3, "Format", f"  Slide {idx+1} [{section_type}]: kept")

    _emit(3, "Format", f"Enriched {enriched_count}/{len(template_slides)} slides")

    # ── Step 4b: Empty-deck guardrail ──
    # If less than 30% of slides got enriched (typical when the chosen template
    # is mostly visual — e.g. Oil & Gas), run a structural fill pass that uses
    # the template's TOC + same-sector KB to GENERATE content for slides where
    # the first pass returned "keep". This makes sure the user always gets a
    # populated deck even when the source template was light on text.
    density = enriched_count / max(1, len(template_slides))
    guardrail_filled = 0
    if density < 0.30:
        _emit(
            3, "Format",
            f"Density {density:.0%} below 30% — running structural fill pass..."
        )
        # Pull a wider sector context for the synthesis prompts
        sector_pool = kb.get_closest_sector_content(
            rfp_sector=rfp_sector, section_type="experience",
            query_text=f"{rfp_sector} engagements credentials methodology", n_results=6,
        )
        sector_tocs = kb.get_sector_tocs(
            rfp_sector, exclude_filename=os.path.basename(template_path)
        )

        for slide in template_slides:
            idx = slide["index"]
            if idx in content_map:
                continue  # already enriched
            purpose = toc_purpose_by_idx.get(idx, "")
            # Skip pure-visual / chapter-marker slides — they don't have text
            # shapes to populate without destroying the design.
            if purpose in ("divider", "closing", "section_header"):
                continue
            if _is_section_header_slide(slide):
                continue
            # Need at least one replaceable text shape
            if not any(s.get("name") for s in slide.get("shapes", []) if s.get("name")):
                continue

            try:
                result = _structural_fill_agent(
                    client, deployment, slide,
                    purpose=purpose or "content",
                    section_type=_normalize_purpose(purpose) if purpose else "other",
                    rfp_summary_text=rfp_summary_text,
                    kb_content=sector_pool,
                    sector_tocs=sector_tocs,
                    client_name=client_name,
                    project_name=project_name,
                    rfp_sector=rfp_sector,
                    user_prompt=user_prompt,
                )
            except Exception as exc:
                _emit(3, "Format", f"  Slide {idx+1}: structural fill error: {exc}")
                continue

            if result and result.get("action") == "modify" and result.get("shapes_content"):
                content_map[idx] = result["shapes_content"]
                enriched_count += 1
                guardrail_filled += 1
                _emit(
                    3, "Format",
                    f"  Slide {idx+1} [{purpose or 'content'}]: FILLED (guardrail)"
                )

        new_density = enriched_count / max(1, len(template_slides))
        _emit(
            3, "Format",
            f"Guardrail: +{guardrail_filled} slides filled "
            f"(density {density:.0%} -> {new_density:.0%})"
        )

    # ── Step 5: Build PPTX ──
    _emit(4, "QA", "Building presentation...")
    placeholders = {
        "[CLIENT]": client_name, "[client]": client_name, "[Client]": client_name,
        "[CLIENT NAME]": client_name, "[Client Name]": client_name, "[client name]": client_name,
        "[COMPANY]": "EY", "[Company]": "EY", "[company]": "EY",
        "[COMAPNY]": "EY", "[Comapny]": "EY",  # known template typo
        "[NAME]": client_name, "[name]": client_name, "[Name]": client_name,
        "[ADDRESS]": kf.get("country_region", ""), "[address]": kf.get("country_region", ""),
        "[PHONE]": "", "[phone]": "",
        "[DATE]": today, "[date]": today, "[Date]": today,
        "[PROJECT]": project_name, "[project]": project_name, "[Project]": project_name,
        "[PROJECT NAME]": project_name, "[Project Name]": project_name,
        "[TENDER]": kf.get("tender_number", "") or "", "[Tender]": kf.get("tender_number", "") or "",
        "[SECTOR]": rfp_sector, "[sector]": rfp_sector, "[Sector]": rfp_sector,
    }
    output_path = generate_pptx_from_enrichments(
        template_path, output_dir, content_map, placeholders,
    )

    # ── Step 6: Verification pass ──
    _emit(5, "Verify", "Running verifier agent...")

    # Gather KB entries across section types so verifier can cross-check named entities
    seen_ids = set()
    all_kb_entries = []
    for section in ("experience", "firm_profile", "team", "methodology", "understanding"):
        try:
            entries = kb.get_closest_sector_content(
                rfp_sector=rfp_sector,
                section_type=section,
                query_text=f"{rfp_sector} {section}",
                n_results=4,
            )
        except Exception:
            entries = []
        for e in entries:
            key = (e.get("template", ""), e.get("title", ""), e.get("content", "")[:80])
            if key in seen_ids:
                continue
            seen_ids.add(key)
            all_kb_entries.append(e)
    if not all_kb_entries:
        all_kb_entries = kb.get_closest_sector_content(
            rfp_sector=rfp_sector, section_type="experience",
            query_text=rfp_sector, n_results=8,
        )

    verify_report = verify_and_fix(
        pptx_path=output_path,
        client=client,
        deployment=deployment,
        rfp_summary_text=rfp_summary_text,
        kb_entries=all_kb_entries,
        client_name=client_name,
        rfp_sector=rfp_sector,
    )
    _emit(5, "Verify", f"Layout fixes: {verify_report['layout_fixes']}, Content fixes: {verify_report['content_fixes']}, Slides deleted: {verify_report['slides_deleted']}")

    elapsed = time.time() - t0
    _emit(6, "Build", f"Done in {elapsed:.1f}s -> {output_path}")

    # Re-read final slide count after curation
    final_prs = Presentation(output_path)
    final_slides = _extract_template_slides_from_prs(final_prs)

    return {
        "output_path": output_path,
        "slides": _build_slide_outline_from_list(final_slides),
        "qa_report": {**_compute_qa(content_map, template_slides), **verify_report},
        "timing_seconds": round(elapsed, 1),
        "sections_mapped": enriched_count,
        "rfp_sections_found": len(rfp_analysis["knowledge_base"]),
    }


# ═══════════════════════════════════════════════════════════════════════════
# THE SLIDE AGENT — per-slide LLM call with KB retrieval + RFP summary
# ═══════════════════════════════════════════════════════════════════════════

SLIDE_AGENT_SYSTEM = """You are a proposal slide editor. You receive:
1. ONE slide from a proposal template (defines the VISUAL FORMAT to match)
2. An RFP SUMMARY (defines what the CLIENT needs — use for understanding/scope/methodology slides)
3. KNOWLEDGE BASE entries (REAL firm content from past proposals — use for experience/team/credentials slides)

CRITICAL: The template may be from a DIFFERENT SECTOR than the RFP. ALL content must be adapted to the RFP's sector.

HARD LENGTH RULES (apply to EVERY slide):
- EVERY shape has a `max_words` field derived from its physical size on the slide. NEVER exceed `max_words` for any shape. This is a HARD CAP because the text physically will not fit otherwise and will overflow other elements.
- Shapes with `max_words < 10` are LABEL shapes (icons captions, pill-buttons, tags). Their text MUST be a short noun phrase of 2-5 words. Do NOT write sentences in label shapes. Do NOT write paragraphs in label shapes.
- Shapes with `max_words < 25` are SUB-HEADING shapes. Keep them to one short sentence.
- Only shapes with `max_words >= 60` (the large body / content placeholders) may hold a paragraph.
- TITLE / HEADING / SUBTITLE SHAPES (any shape whose `is_title` flag is true, or whose name contains "Title", "Heading", or "Subtitle"):
  - HARD CAP: max 2 lines, max 12 words, max 80 characters AND <= the shape's `max_words` budget.
  - The original shape's word_count is your ceiling — never exceed it by more than +20%.
  - If you need to convey a long RFP section name, abbreviate it (e.g. "Site Visit and Verification of Information" -> "Site Visit & Verification"; "Submission and Uploading of Bids" -> "Bid Submission").
  - NEVER append descriptive sentences to a title. Descriptions belong in body shapes only.
- BODY SHAPES: stay within original `word_count` +/- 30% AND <= the shape's `max_words` budget. If the two limits disagree, the SMALLER one wins.
- For shapes that are clearly the cover-slide title (first slide title), the text must be the CLIENT NAME alone (≤ 6 words). Do not append the project description there.
- DO NOT WRITE PARAGRAPHS in any shape whose `max_words` is < 30. Use a short phrase or a single concise sentence.
- DO NOT NARRATE — if you cannot fit a methodology step inside its shape, choose fewer words; never let text spill past the budget. Truncation downstream will visibly cut content if you exceed the cap.

ALIGNMENT / VISUAL CONSISTENCY:
- The slide already defines its own visual rhythm (icon row, left/right column, centered title, etc.). Match the EXISTING shape's content scope — if a shape was a one-line label in the template, keep your replacement a one-line label. If it was a 3-bullet box, write 3 short bullets. Do not change layout intent.

SECTION-SPECIFIC RULES:

FOR "experience" / "firm_profile" / "team" SLIDES:
- Use the KNOWLEDGE BASE entries as your source. These are REAL case studies and credentials.
- Adapt them to highlight relevance to the RFP's sector, but keep the core facts real.
- NEVER fabricate projects, case studies, or client engagements.
- ABSOLUTELY NEVER put the RFP client's name (the name from CLIENT field above) into experience or case study descriptions. The RFP client is a PROSPECTIVE client — we have NOT worked with them yet.
- HARD RULE — NAMED ENTITIES: Do NOT introduce any specific organization name, acronym, brand, program name, or proper-noun client reference unless that exact name appears verbatim in the KB entries below. If you need to refer to a past client and the KB doesn't name one, use a generic description ("a leading government entity", "a major infrastructure authority", "a national development agency"). Inventing names like "LEHS", "ACME Health", or any other acronym is forbidden.
- If KB has a healthcare case study, describe the same TYPE of advisory work but generalize it for the RFP's sector.
- Keep specific details from KB entries (team sizes, durations, methodologies used) — just adapt the sector context. Do NOT invent numbers or dates that are not in the KB or RFP.

FOR "cover_letter" SLIDES (section_type == "cover_letter", or template text contains "Dear Sir" / "Thank you for inviting"):
- Treat the largest body shape as a FORMAL BUSINESS LETTER. The KB cover_letter entries are the canonical structure — mirror their format.
- Required paragraph structure (one paragraph per intent, separated by blank lines):
  1. Greeting line ("Dear Sir," / "Dear [Mr./Ms. Last Name],"). Use the original greeting if present.
  2. Opening paragraph — thank the client by name (full RFP CLIENT name, not a placeholder) for the RFI invitation, name the project from the RFP summary, and confirm EY's intent to support it.
  3. Why-us paragraph — EY's sector credentials drawn from the KB (generalized; no fabricated org names). Mention consortium/partner only if KB cover-letter entries mention one.
  4. RFI-scope paragraph — clarify that this response is for discussion purposes pending full pricing/risk/contract alignment (only if the original template has this paragraph).
  5. Closing paragraph — willingness to meet, contact statement, signed off by EY.
- Preserve all paragraph breaks. NEVER collapse the letter into one paragraph.
- Length: stay within +/- 20% of the original letter body word count.
- The smaller "address block" shape (containing "To:", "[NAME]", "[CLIENT]", "[ADDRESS]", tender details) MUST keep its line-by-line structure. Replace each bracketed line individually:
   To:  <addressee from RFP (or "The Tender Committee" if unknown)>
   <Department or division from RFP (or generic "Procurement Office")>
   <CLIENT NAME>
   <Country / Region from RFP>
   Telephone: <leave blank or use RFP-provided contact>
   Tender Ref: <Tender number from RFP, or keep XXXXXXXXXX if unknown>
   Tender Name: <Tender / project name from RFP>
- The small date shape must contain today's date in "DD Month YYYY" format.
- NEVER leave any [BRACKET] tokens behind in cover-letter shapes.

FOR "understanding" / "methodology" / "timeline" / "pricing" / "cover" SLIDES:
- Use the RFP SUMMARY as your source. Generate content specific to THIS RFP's requirements.
- Reference specific RFP facts: dates, evaluation criteria, deliverables, scope items.
- For COVER (title) slides specifically: the main title shape must be JUST the CLIENT NAME (no descriptions). The subtitle/tagline shape may carry the project / tender name in <= 10 words.

FOR ALL SLIDES:
- Match the original slide's FORMAT (bullets vs paragraphs, word count +/- 30%)
- NEVER use placeholder brackets like [CLIENT], [COMPANY] — use actual names provided
- Professional tone for executive presentations

Return JSON:
{
  "action": "keep" or "modify",
  "reason": "one-line reason",
  "shapes_content": {
    "Shape Name 1": "New text with \\n for line breaks",
    "Shape Name 2": "KEEP"
  }
}

Write "KEEP" for: shapes with <5 words, slide numbers, footers, dates, freeforms, connectors, pictures, groups.
Write "keep" as action ONLY for: blank slides, pure dividers with <5 words, TOC/agenda slides."""


def _slide_agent(
    client, deployment, slide,
    rfp_summary_text, kb_content, section_type,
    client_name, project_name, rfp_sector, user_prompt,
    slide_purpose="",
    today_date="", tender_number="", country_region="",
):
    shapes_context = []
    for s in slide["shapes"]:
        if not s.get("name"):
            continue
        wc = s.get("word_count", 0)
        # Skip decorative empty shapes. Many templates have empty Rectangle /
        # Oval / Freeform shapes layered as visual backgrounds (e.g. the 6
        # "Rectangle 60" highlight boxes behind the actual content rectangles
        # on Healthcare slide 11). If we feed them to the agent it will
        # generate text into them and we get layered / duplicated content on
        # top of the real content shapes. We ONLY expose empty shapes to the
        # agent when their name signals a true content placeholder.
        if wc < 2 and not _is_named_content_placeholder(s.get("name", "")):
            continue
        shapes_context.append({
            "name": s["name"],
            "text": (s.get("text") or "")[:400],
            "word_count": wc,
            "is_title": s.get("is_title", False),
            "size_in": f"{s.get('w_in', 0)}x{s.get('h_in', 0)}",
            "max_words": s.get("max_words", 30),
        })
    if not shapes_context:
        return {"action": "keep", "reason": "no shapes"}

    # Format KB entries for the prompt
    kb_text = ""
    if kb_content:
        kb_text = "\n\nKNOWLEDGE BASE (REAL firm content — use as source for experience/team slides):\n"
        for i, entry in enumerate(kb_content, 1):
            kb_text += f"\n--- KB Entry {i} (from {entry['template']}, {entry['sector']}, {entry['section_type']}) ---\n"
            kb_text += entry["content"][:500] + "\n"

    user_msg = (
        f"CLIENT: {client_name}\n"
        f"PROJECT: {project_name}\n"
        f"RFP SECTOR: {rfp_sector}\n"
        f"COUNTRY/REGION: {country_region}\n"
        f"TENDER NUMBER: {tender_number}\n"
        f"TODAY'S DATE: {today_date}\n"
        f"THIS SLIDE'S SECTION TYPE: {section_type}\n"
        f"TOC PURPOSE: {slide_purpose or section_type}\n\n"
        f"SLIDE {slide['index']+1} (layout: {slide['layout']}):\n"
        f"Shapes:\n{json.dumps(shapes_context, indent=2)}\n\n"
        f"RFP SUMMARY (use for understanding/scope/methodology content):\n{rfp_summary_text}\n"
        f"{kb_text}"
    )
    if user_prompt:
        user_msg += f"\n\nUSER INSTRUCTIONS: {user_prompt}"

    try:
        resp = client.chat.completions.create(
            model=deployment,
            messages=[
                {"role": "system", "content": SLIDE_AGENT_SYSTEM},
                {"role": "user", "content": user_msg[:7000]},  # stay within token limits
            ],
            response_format={"type": "json_object"},
            temperature=0.1,
            max_tokens=4096,
        )
        result = json.loads(resp.choices[0].message.content)

        # Filter "KEEP" entries
        if result.get("action") == "modify" and result.get("shapes_content"):
            filtered = {k: v for k, v in result["shapes_content"].items()
                        if isinstance(v, str) and v.strip().upper() != "KEEP"}
            if not filtered:
                return {"action": "keep", "reason": "all shapes kept"}
            result["shapes_content"] = filtered

        return result
    except Exception as e:
        print(f"      Agent error slide {slide['index']+1}: {e}")
        return {"action": "keep", "reason": f"error: {e}"}


# ═══════════════════════════════════════════════════════════════════════════
# Helpers
# ═══════════════════════════════════════════════════════════════════════════

def _format_rfp_summary(summary: dict) -> str:
    parts = []
    for key in ["one_liner", "client", "project", "sector", "timeline", "budget",
                "methodology_expectations", "team_expectations"]:
        val = summary.get(key)
        if val:
            parts.append(f"{key.upper().replace('_', ' ')}: {val}")
    for key in ["scope", "deliverables", "evaluation_criteria", "key_requirements"]:
        val = summary.get(key)
        if val and isinstance(val, list):
            parts.append(f"{key.upper().replace('_', ' ')}:")
            for item in val:
                parts.append(f"  - {item}")
        elif val:
            parts.append(f"{key.upper().replace('_', ' ')}: {val}")
    return "\n".join(parts)


_PLACEHOLDER_NAME_RE = re.compile(
    r"^\s*("
    r"title|subtitle|content placeholder|body placeholder|text placeholder|"
    r"picture placeholder|placeholder"
    r")\b",
    re.IGNORECASE,
)


def _is_named_content_placeholder(name: str) -> bool:
    """True if a shape's name is one of the standard PowerPoint placeholder
    names (Title, Subtitle, Content Placeholder, etc.). These represent real
    fillable regions, even when empty. All other empty shapes are treated as
    decorative (background rectangles, layered highlight boxes, ovals, etc.).
    """
    if not name:
        return False
    return bool(_PLACEHOLDER_NAME_RE.match(name.strip()))


def _is_section_header_slide(slide: dict) -> bool:
    """Heuristic: a section starter slide has a 1-2 digit number shape plus a
    short title (2-7 words), and total content text is <= ~15 words.
    Example: "01 / Why Us and Our Point of View" with an empty body rectangle.
    These should never be enriched — the agent would otherwise dump a paragraph
    into the empty body, destroying the design.

    Care is taken NOT to flag generic numbered footers (page numbers) on
    otherwise-content slides. We require BOTH a prominent number AND a
    multi-word title; a slide with only "3" and no clear title is not a header.
    """
    shapes = slide.get("shapes", []) or []
    if not shapes:
        return False
    total_words = sum(s.get("word_count", 0) for s in shapes)
    if total_words > 15 or total_words < 2:
        return False
    number_count = 0
    title_count = 0
    for s in shapes:
        text = (s.get("text") or "").strip()
        if not text:
            continue
        wc = s.get("word_count", 0)
        # Section number like "01", "02", "1.", "I", "II"
        cleaned = text.replace(".", "").strip()
        if 1 <= len(text) <= 4 and cleaned.isdigit():
            number_count += 1
            continue
        # Heading-style title:
        #   * 2-7 words (e.g. "Why Us and Our Point of View")  OR
        #   * 1 substantive word with at least 6 chars (e.g. "Methodology")
        if 2 <= wc <= 7 and not text.endswith(".") and not text.endswith(","):
            title_count += 1
        elif wc == 1 and len(text) >= 6 and not text.endswith(".") and text[0].isalpha():
            title_count += 1
    return number_count == 1 and title_count == 1


def _extract_template_slides(pptx_path):
    prs = Presentation(pptx_path)
    slides = []
    for i, slide in enumerate(prs.slides):
        shapes_data = []
        all_text = []
        title_text = ""
        name_counts = {}  # track duplicate names
        for shape in slide.shapes:
            try:
                w_in = round((shape.width or 0) / 914400, 2)
                h_in = round((shape.height or 0) / 914400, 2)
            except Exception:
                w_in = h_in = 0.0
            # Rough text-budget heuristic based on shape area. Tunes the agent
            # so it doesn't try to inject paragraphs into label-sized boxes.
            # Empirically a label shape (2x0.5 in) fits ~6 words; a content
            # placeholder (6x4 in) fits ~120 words.
            area = max(w_in * h_in, 0.05)
            max_words_for_shape = max(3, int(area * 7))
            if shape.has_text_frame:
                text = shape.text_frame.text or ""
                wc = len(text.split())
                # Disambiguate duplicate shape names
                raw_name = shape.name
                name_counts[raw_name] = name_counts.get(raw_name, 0) + 1
                unique_name = raw_name if name_counts[raw_name] == 1 else f"{raw_name} #{name_counts[raw_name]}"
                shapes_data.append({
                    "name": unique_name, "text": text[:500],
                    "word_count": wc, "is_title": "title" in raw_name.lower(),
                    "w_in": w_in, "h_in": h_in,
                    "max_words": max_words_for_shape,
                })
                if text.strip():
                    all_text.append(text.strip())
                if "title" in raw_name.lower() and text.strip():
                    title_text = text.strip()[:100]
            elif shape.has_table:
                shapes_data.append({
                    "name": shape.name, "text": "(table)",
                    "word_count": 0, "is_title": False,
                    "w_in": w_in, "h_in": h_in,
                    "max_words": max_words_for_shape,
                })
        if not title_text and all_text:
            title_text = all_text[0][:100]
        slides.append({
            "index": i, "layout": slide.slide_layout.name,
            "title": title_text, "shapes": shapes_data,
            "full_text": "\n".join(all_text)[:1000],
            "total_words": sum(s["word_count"] for s in shapes_data),
        })
    return slides


def _extract_template_slides_from_prs(prs):
    """Extract slide info from an already-open Presentation object."""
    slides = []
    for i, slide in enumerate(prs.slides):
        all_text = []
        title_text = ""
        for shape in slide.shapes:
            if shape.has_text_frame:
                text = shape.text_frame.text or ""
                if text.strip():
                    all_text.append(text.strip())
                if "title" in shape.name.lower() and text.strip():
                    title_text = text.strip()[:100]
        if not title_text and all_text:
            title_text = all_text[0][:100]
        slides.append({"index": i, "title": title_text})
    return slides


def _build_slide_outline_from_list(slides_list):
    return [
        {
            "slide_number": s["index"] + 1,
            "title": s["title"][:80] or f"Slide {s['index']+1}",
            "slide_type": "unknown",
        }
        for s in slides_list
    ]


def _build_slide_outline(template_slides, content_map):
    return [
        {
            "slide_number": s["index"] + 1,
            "title": s["title"][:80] or f"Slide {s['index']+1}",
            "slide_type": "rfp_content" if s["index"] in content_map else "original",
        }
        for s in template_slides
    ]


def _compute_qa(content_map, template_slides):
    enriched = len(content_map)
    total = len(template_slides)
    return {
        "overall_score": min(100, int(50 + (enriched / max(total, 1)) * 100)),
        "verdict": f"{enriched} slides enriched, {total - enriched} kept original",
        "slides_enriched": enriched,
        "slides_kept_original": total - enriched,
    }


# ═══════════════════════════════════════════════════════════════════════════
# Structural-fill agent (Pass B for the empty-deck guardrail)
# ═══════════════════════════════════════════════════════════════════════════

_PURPOSE_ALIASES = {
    "title_slide": "cover",
    "section_header": "other",
    "diagram": "methodology",
    "table": "pricing",
    "cv": "team",
    "closing": "other",
    "divider": "other",
    "exec_summary": "understanding",
    "toc": "toc",
    "content": "other",
}


def _normalize_purpose(purpose: str) -> str:
    if not purpose:
        return "other"
    p = purpose.lower()
    if p in (
        "cover", "cover_letter", "toc", "firm_profile", "experience",
        "understanding", "methodology", "team", "timeline", "pricing", "appendix", "other",
    ):
        return p
    return _PURPOSE_ALIASES.get(p, "other")


STRUCTURAL_FILL_SYSTEM = """You are designing a proposal slide FROM SCRATCH.

The slide currently has no meaningful text (the template ships it empty or
just visually). You will FILL its shapes with new content that fits:

1. The slide's TOC PURPOSE (cover / understanding / methodology / experience /
   team / timeline / pricing / firm_profile / appendix / other).
2. The RFP SUMMARY (what the client needs and the engagement scope).
3. The KNOWLEDGE BASE entries (real EY credentials and past work — generalize
   the sector but never invent named clients or projects).
4. The SECTOR TOCs (how other same-sector proposals structured this section).

HARD RULES:
- Every shape has a `max_words` field derived from its physical size. NEVER
  exceed it. Smaller box -> shorter text. Label-sized boxes (max_words < 10)
  hold ONLY a 2-5 word phrase — never a sentence, never a paragraph.
- TITLE shapes (name contains "Title" / "Heading"): max 12 words, max 2 lines,
  AND <= shape's `max_words`.
- BODY / CONTENT shapes (placeholder, body, text) with `max_words >= 60`:
  30-90 words, prefer 3-5 short bullets joined with \\n.
- SUBTITLE shapes: max 14 words AND <= `max_words`.
- NEVER invent client names, project names, or organization names that are not
  in the RFP or KB entries.
- NEVER leave [BRACKET] tokens behind.
- For purposes like "divider" / "closing" / pure cover, return "keep".

Return JSON:
{
  "action": "modify" or "keep",
  "reason": "<one line>",
  "shapes_content": {
    "<exact shape name>": "<new text>",
    ...
  }
}

Write "KEEP" as the value for shapes that are decorative (freeforms,
connectors, images, footers, page numbers, group shapes)."""


def _structural_fill_agent(
    client, deployment, slide,
    purpose, section_type,
    rfp_summary_text, kb_content, sector_tocs,
    client_name, project_name, rfp_sector, user_prompt,
):
    """LLM call that GENERATES content for an otherwise-empty slide based on
    its TOC purpose + RFP + same-sector knowledge."""
    shapes_context = []
    for s in slide.get("shapes", []):
        if not s.get("name"):
            continue
        wc = s.get("word_count", 0)
        # Same decorative-shape filter as Pass A: only fill true content
        # placeholders (Title / Subtitle / Content Placeholder / Body
        # Placeholder / Text Placeholder) when the slide is otherwise empty.
        # This prevents the structural-fill agent from writing into every
        # decorative Rectangle / Oval / Freeform.
        if wc < 2 and not _is_named_content_placeholder(s.get("name", "")):
            continue
        shapes_context.append({
            "name": s["name"],
            "text": (s.get("text") or "")[:200],
            "word_count": wc,
            "is_title": s.get("is_title", False),
            "size_in": f"{s.get('w_in', 0)}x{s.get('h_in', 0)}",
            "max_words": s.get("max_words", 30),
        })
    if not shapes_context:
        return {"action": "keep", "reason": "no named shapes"}

    kb_text = ""
    if kb_content:
        kb_text = "\n\nKNOWLEDGE BASE (real firm content — generalize for this sector):\n"
        for i, entry in enumerate(kb_content[:5], 1):
            kb_text += (
                f"\n--- KB Entry {i} ({entry.get('template','')} / "
                f"{entry.get('sector','')} / {entry.get('section_type','')}) ---\n"
                + (entry.get("content", "") or "")[:400] + "\n"
            )

    toc_text = ""
    if sector_tocs:
        toc_text = "\n\nSECTOR TOCs (structural hints from other same-sector decks):\n"
        for toc in sector_tocs[:3]:
            sections = toc.get("sections") or []
            toc_text += (
                f"- {toc.get('template','')}: "
                + " | ".join(
                    f"{sec.get('name','')}({sec.get('purpose','')})"
                    for sec in sections[:10]
                )
                + "\n"
            )

    user_msg = (
        f"CLIENT: {client_name}\n"
        f"PROJECT: {project_name}\n"
        f"RFP SECTOR: {rfp_sector}\n"
        f"THIS SLIDE'S TOC PURPOSE: {purpose}\n"
        f"NORMALIZED SECTION TYPE: {section_type}\n\n"
        f"SLIDE {slide['index']+1} (layout: {slide.get('layout','')}):\n"
        f"Shapes (currently empty or sparse):\n"
        f"{json.dumps(shapes_context, indent=2)}\n\n"
        f"RFP SUMMARY:\n{rfp_summary_text}\n"
        f"{kb_text}{toc_text}"
    )
    if user_prompt:
        user_msg += f"\n\nUSER INSTRUCTIONS: {user_prompt}"

    resp = client.chat.completions.create(
        model=deployment,
        messages=[
            {"role": "system", "content": STRUCTURAL_FILL_SYSTEM},
            {"role": "user", "content": user_msg[:7500]},
        ],
        response_format={"type": "json_object"},
        temperature=0.2,
        max_tokens=2048,
    )
    result = json.loads(resp.choices[0].message.content)
    if result.get("action") == "modify" and result.get("shapes_content"):
        filtered = {
            k: v for k, v in result["shapes_content"].items()
            if isinstance(v, str) and v.strip().upper() != "KEEP" and v.strip()
        }
        if not filtered:
            return {"action": "keep", "reason": "all shapes kept"}
        result["shapes_content"] = filtered
    return result
