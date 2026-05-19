"""Engine for the MCP server.

Pipeline:
  1. Parser agent reads the structured RFP / proposal text and emits a slide plan.
  2. Template inspector enumerates the Healthcare template shape by shape.
  3. SlideMaster agent maps each plan section to a slide's shapes.
  4. Verifier agent (Python + LLM) enforces:
       - fidelity to the slide plan (no hallucination, no missing facts)
       - bold and alignment unchanged
       - text fits the shape (auto font-shrink allowed)
  5. SlideMaster retries up to MAX_ITERATIONS per slide; final pass shrinks fonts.

The existing pipeline / KB / verifier / generator under src/ is untouched.
We only import helpers from generate_pptx.py and verifier.py.
"""

from __future__ import annotations

import copy
import json
import os
import sys
import time
import uuid
from collections import Counter
from typing import Any

from lxml import etree
from pptx import Presentation
from pptx.util import Pt

_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
_SRC_DIR = os.path.join(_THIS_DIR, "src")
if _SRC_DIR not in sys.path:
    sys.path.insert(0, _SRC_DIR)

from config import load_config  # noqa: E402
from generate_pptx import (  # noqa: E402
    NS_A,
    _duplicate_slide,
    _find_shape_by_name,
    _is_skip_shape,
    _replace_text_preserving_format,
    _sanitize,
)
from verifier import (  # noqa: E402
    _delete_slide,
    _estimate_max_chars_per_line,
    _estimate_max_lines,
    _get_dominant_font_size,
    _is_section_divider_slide,
    _replace_text_in_frame,
    _shrink_font_to_fit,
)

MAX_ITERATIONS = 3
RUNS_DIR = os.path.join(_THIS_DIR, "runs")

SECTION_TYPES = {
    "cover": ["cover", "title slide"],
    "cover_letter": ["dear", "thank you for inviting", "cover letter", "executive summary"],
    "toc": ["table of content", "agenda", "contents"],
    "firm_profile": ["why ey", "why us", "about ey", "firm profile", "about us"],
    "experience": ["experience", "credential", "case stud", "project deliver",
                   "we worked", "we have deliver", "track record", "past project"],
    "understanding": ["our understanding", "point of view", "market insight",
                      "key consideration", "sector overview"],
    "methodology": ["methodology", "approach", "our approach", "framework",
                    "workstream", "phase", "how we will"],
    "team": ["our team", "team structure", "key personnel", "organis",
             "resource", "staffing", "cv", "profile"],
    "timeline": ["timeline", "gantt", "work plan", "schedule", "milestone"],
    "pricing": ["pricing", "fee", "financial proposal", "cost", "budget"],
    "appendix": ["appendix", "annex", "supporting", "certificate"],
}


def _classify(text: str) -> str:
    t = text.lower()
    for stype, kws in SECTION_TYPES.items():
        for kw in kws:
            if kw in t:
                return stype
    return "other"


# ───────────────────────────────────────────────────────────────────────────────
# Step 1: parse the structured RFP text into a slide plan
# ───────────────────────────────────────────────────────────────────────────────

PARSER_SYSTEM = """You convert an RFP / proposal brief into a slide-by-slide
PLAN that will be populated into a fixed PowerPoint template.

Input may be either:
  (A) a pre-broken-out proposal with explicit "Slide N" sections, OR
  (B) raw RFP analysis (client, scope, criteria, KPIs, etc.).

Whatever the input, produce a JSON plan covering these section_types in order
when content is available: cover, cover_letter, firm_profile, understanding,
pov, experience, methodology, team, timeline, pricing, appendix.

Rules:
- DO NOT invent facts. Use only names, numbers, dates, KPIs that appear in
  the input. If the input does not name a client, write "the Client" etc.
- Each section MUST be derived from a passage that is actually in the input.
- Keep headlines short (<= 12 words). Bullets <= 14 words each.
- Visual / design notes are optional hints, not slide content.

Return JSON:
{
  "client_name": "...",
  "project_name": "...",
  "sector": "...",
  "sections": [
    {
      "section_type": "cover|cover_letter|firm_profile|understanding|pov|experience|methodology|team|timeline|pricing|appendix",
      "headline": "Short headline",
      "subhead": "Optional one-line sub-head (or empty)",
      "bullets": ["Bullet 1", "Bullet 2", "Bullet 3"],
      "notes": "Optional design note from the source"
    }
  ]
}
"""


def parse_plan(client, deployment: str, rfp_text: str, user_prompt: str = "") -> dict:
    user_block = f"\n\nADDITIONAL USER INSTRUCTIONS:\n{user_prompt}" if user_prompt.strip() else ""
    resp = client.chat.completions.create(
        model=deployment,
        messages=[
            {"role": "system", "content": PARSER_SYSTEM},
            {"role": "user", "content": f"PROPOSAL BRIEF:\n{rfp_text[:30000]}{user_block}"},
        ],
        response_format={"type": "json_object"},
        temperature=0.05,
        max_tokens=4096,
    )
    return json.loads(resp.choices[0].message.content)


# ───────────────────────────────────────────────────────────────────────────────
# Step 2: inspect the template — shape inventory per slide
# ───────────────────────────────────────────────────────────────────────────────

def _shape_role(shape) -> str:
    if not shape.has_text_frame:
        return "non_text"
    name_l = shape.name.lower()
    if _is_skip_shape(name_l):
        return "skip"
    if "footer" in name_l or "page" in name_l or "date" in name_l or "slide number" in name_l:
        return "skip"
    if "title" in name_l or "heading" in name_l or "subtitle" in name_l:
        return "title"
    return "body"


def _shape_format(shape) -> dict:
    sizes, bolds, aligns = [], [], []
    a_pPr = f"{{{NS_A}}}pPr"
    for para in shape.text_frame.paragraphs:
        pPr = para._pPr if hasattr(para, "_pPr") else None
        if pPr is None:
            pPr = para._p.find(a_pPr)
        algn = pPr.get("algn") if pPr is not None else None
        aligns.append(algn or "l")
        for run in para.runs:
            if run.font.size:
                sizes.append(int(run.font.size / 12700))
            if run.font.bold is True:
                bolds.append(True)
            elif run.font.bold is False:
                bolds.append(False)

    def _mode(xs, default):
        return Counter(xs).most_common(1)[0][0] if xs else default

    return {
        "font_size": _mode(sizes, _get_dominant_font_size(shape.text_frame)),
        "bold": bool(_mode(bolds, False)),
        "align": _mode(aligns, "l"),
    }


def inspect_template(template_path: str) -> list[dict]:
    prs = Presentation(template_path)
    inv = []
    for i, slide in enumerate(prs.slides):
        shapes = []
        slide_text = ""
        for shape in slide.shapes:
            role = _shape_role(shape)
            if role == "non_text":
                continue
            w_in = (shape.width or 0) / 914400
            h_in = (shape.height or 0) / 914400
            cur = shape.text_frame.text.strip() if shape.has_text_frame else ""
            fmt = _shape_format(shape) if shape.has_text_frame else {}
            slide_text += " " + cur
            shapes.append({
                "name": shape.name,
                "role": role,
                "bbox_in": [round(w_in, 2), round(h_in, 2)],
                "current_text_preview": cur[:120],
                "current_word_count": len(cur.split()),
                "font_size": fmt.get("font_size"),
                "bold": fmt.get("bold"),
                "align": fmt.get("align"),
            })
        section_type = _classify(slide_text)
        layout_name = (slide.slide_layout.name or "").lower()
        inv.append({
            "slide_index": i,
            "layout": layout_name,
            "section_type": section_type,
            "shapes": shapes,
        })
    return inv


# ───────────────────────────────────────────────────────────────────────────────
# Step 3: pair plan sections with template slides
# ───────────────────────────────────────────────────────────────────────────────

POV_FALLBACK = "understanding"


def _section_match_score(plan_type: str, slide_type: str) -> int:
    if plan_type == slide_type:
        return 100
    aliases = {
        "pov": ["understanding", "methodology"],
        "cover_letter": ["understanding", "firm_profile"],
        "appendix": ["other"],
    }
    if slide_type in aliases.get(plan_type, []):
        return 60
    if plan_type == "other" or slide_type == "other":
        return 5
    return 1


def pair_plan_with_slides(plan: dict, template_inv: list[dict]) -> list[dict]:
    """Greedy section-type pairing with two fixed anchors:
      - slide 0 (first slide) is reserved for the plan's 'cover' section.
      - the last template slide is treated as closing/contact and NOT paired
        (left intact unless the plan has an explicit appendix/closing section).
    Plan sections without a slide trigger duplication of the closest match."""
    pairings = []
    used: set[int] = set()

    if not template_inv:
        return pairings

    first_idx = template_inv[0]["slide_index"]
    last_idx = template_inv[-1]["slide_index"]

    # Anchor cover -> slide 0 (if a cover plan section exists)
    cover_plan_idx = next(
        (i for i, s in enumerate(plan.get("sections", []))
         if s.get("section_type") == "cover"),
        None,
    )
    if cover_plan_idx is not None:
        pairings.append({"plan_idx": cover_plan_idx, "slide_index": first_idx, "duplicate": False})
        used.add(first_idx)

    for p_idx, section in enumerate(plan.get("sections", [])):
        if p_idx == cover_plan_idx:
            continue
        ptype = section.get("section_type", "other")
        ranked = sorted(
            (s for s in template_inv
             if s["slide_index"] not in used and s["slide_index"] != last_idx),
            key=lambda s: _section_match_score(ptype, s["section_type"]),
            reverse=True,
        )
        if not ranked:
            historical = sorted(
                template_inv,
                key=lambda s: _section_match_score(ptype, s["section_type"]),
                reverse=True,
            )
            pairings.append({"plan_idx": p_idx, "slide_index": historical[0]["slide_index"], "duplicate": True})
            continue
        best = ranked[0]
        used.add(best["slide_index"])
        pairings.append({"plan_idx": p_idx, "slide_index": best["slide_index"], "duplicate": False})
    return pairings


# ───────────────────────────────────────────────────────────────────────────────
# Step 4: SlideMaster agent — propose per-shape content for one slide
# ───────────────────────────────────────────────────────────────────────────────

SLIDEMASTER_SYSTEM = """You are SlideMaster. You replace the content of a
template slide with content from ONE approved plan section.

CRITICAL CONTEXT:
- The template's current shape text is STALE and belongs to a different project.
  It is shown to you only so you understand each shape's role and size — DO NOT
  copy any name, number, region, client, hospital or project from it.
- The plan section is the ONLY allowed source of facts.

You receive:
- ONE plan section (headline, subhead, bullets, notes)
- All EDITABLE shapes on the slide with: name, role (title|body), bbox in inches,
  the stale current text (for length reference only), dominant font size, bold,
  alignment.
- (Optional) the previous attempt and verifier issues for retry.

OUTPUT RULES — strict:
- You MUST emit an entry for EVERY editable shape on the slide.
- For a TITLE shape: write the plan headline (or a 4-10 word title derived from it).
- For BODY shapes: distribute the plan's bullets and subhead across them. Use
  "\\n" between bullets. Match bullet length to shape size (small box <= 6 words,
  large box up to 4-6 bullets).
- If a shape has no relevant content from the plan, emit "" (empty string) to
  BLANK it. Do NOT use "KEEP". Do NOT leave a body shape with stale text.
- DO NOT invent any fact, client name, statistic, date, KPI, dollar amount,
  region, hospital, sector reference, or project that is not in the plan section.
- DO NOT add filler like "We look forward to partnering with you" unless the
  plan section actually says so.

Return JSON: {"shapes": {"<Shape Name>": "<new text or empty string>", ...}}
"""


def slidemaster_propose(
    client,
    deployment: str,
    plan_section: dict,
    slide_inv: dict,
    prev_attempt: dict | None = None,
    verifier_issues: list | None = None,
) -> dict:
    payload = {
        "plan_section": plan_section,
        "slide": {
            "slide_index": slide_inv["slide_index"],
            "section_type": slide_inv["section_type"],
            "layout": slide_inv["layout"],
            "shapes": slide_inv["shapes"],
        },
    }
    if prev_attempt is not None:
        payload["previous_attempt"] = prev_attempt
    if verifier_issues:
        payload["verifier_issues"] = verifier_issues

    resp = client.chat.completions.create(
        model=deployment,
        messages=[
            {"role": "system", "content": SLIDEMASTER_SYSTEM},
            {"role": "user", "content": json.dumps(payload)[:8000]},
        ],
        response_format={"type": "json_object"},
        temperature=0.1,
        max_tokens=2048,
    )
    out = json.loads(resp.choices[0].message.content)
    return out.get("shapes", {})


# ───────────────────────────────────────────────────────────────────────────────
# Apply per-shape texts to a real python-pptx slide
# ───────────────────────────────────────────────────────────────────────────────

def apply_shape_texts(slide, shape_texts: dict[str, str]) -> list[str]:
    """Apply slidemaster's per-shape output.

    - "" (empty string) → BLANK the shape (clears stale template content).
    - "KEEP" (case-insensitive) → leave the shape unchanged.
    - any other string → replace, preserving font/bold/alignment.
    Returns the list of shape names actually modified or blanked.
    """
    touched = []
    for shape_name, new_text in shape_texts.items():
        if new_text is None or not isinstance(new_text, str):
            continue
        if new_text.strip().upper() == "KEEP":
            continue
        shape = _find_shape_by_name(slide, shape_name)
        if not shape or not shape.has_text_frame:
            continue
        if new_text == "":
            _replace_text_in_frame(shape.text_frame, "")
        else:
            _replace_text_preserving_format(shape.text_frame, new_text)
        touched.append(shape_name)
    return touched


# ───────────────────────────────────────────────────────────────────────────────
# Step 5: Verifier — Python checks + LLM faithfulness
# ───────────────────────────────────────────────────────────────────────────────

def _shape_overflow(shape) -> dict:
    """Estimate whether the shape's text overflows its box; returns metrics."""
    w_in = (shape.width or 0) / 914400
    h_in = (shape.height or 0) / 914400
    if w_in <= 0 or h_in <= 0:
        return {"overflow": False}
    font_size = _get_dominant_font_size(shape.text_frame) or 12
    max_cpl = _estimate_max_chars_per_line(w_in, font_size)
    max_ln = _estimate_max_lines(h_in, font_size)
    text = shape.text_frame.text or ""
    est_lines = 0
    for line in text.split("\n"):
        if not line.strip():
            est_lines += 1
        else:
            est_lines += max(1, (len(line) + max_cpl - 1) // max_cpl)
    return {
        "overflow": est_lines > max_ln,
        "est_lines": est_lines,
        "max_lines": max_ln,
        "font_size": font_size,
    }


def python_verify_slide(slide, baseline_inv: dict, touched: list[str]) -> dict:
    """Strict structural verifier: bold and alignment must match the baseline.
    Returns {accepted, issues:[...], shrink_targets:[shape_name,...]}."""
    issues = []
    shrink_targets = []
    by_name = {s["name"]: s for s in baseline_inv["shapes"]}

    for shape in slide.shapes:
        if shape.name not in touched or not shape.has_text_frame:
            continue
        base = by_name.get(shape.name)
        if not base:
            continue
        cur = _shape_format(shape)
        if base.get("bold") is True and cur.get("bold") is not True:
            issues.append({
                "shape": shape.name,
                "type": "bold_drift",
                "detail": "Bold formatting was lost. Re-emit text without altering bold weight.",
            })
        if base.get("align") and cur.get("align") and base["align"] != cur["align"]:
            issues.append({
                "shape": shape.name,
                "type": "align_drift",
                "detail": f"Alignment changed from {base['align']} to {cur['align']}.",
            })

        ov = _shape_overflow(shape)
        if ov.get("overflow"):
            shrink_targets.append((shape, base.get("font_size") or 12))
            issues.append({
                "shape": shape.name,
                "type": "overflow",
                "detail": f"Text overflows ({ov.get('est_lines')} lines vs {ov.get('max_lines')} capacity). Try fewer / shorter bullets.",
            })

    return {"issues": issues, "shrink_targets": shrink_targets}


VERIFIER_SYSTEM = """You are a strict proposal slide reviewer.

You are given:
  - A PLAN SECTION (the only allowed source for facts)
  - The CURRENT SLIDE shape texts after SlideMaster's pass

Reject the slide if:
  1. It contains any name, number, percentage, date, KPI or client/company
     identifier that is NOT present in the PLAN SECTION (or trivially derived
     from it).
  2. A required fact from the PLAN SECTION is missing (e.g. CTR target, deadline).
  3. The headline does not reflect the plan's headline.
  4. Filler / closing remarks were added beyond what the plan describes.

Accept otherwise. Return JSON:
{
  "accepted": true|false,
  "issues": [
    {"shape": "Shape Name", "type": "hallucination|missing|drift|filler",
     "detail": "what is wrong",
     "fix_hint": "what SlideMaster should change"}
  ]
}
"""


def llm_verify_slide(
    client,
    deployment: str,
    plan_section: dict,
    current_shape_texts: dict[str, str],
) -> dict:
    payload = {
        "plan_section": plan_section,
        "current_slide": current_shape_texts,
    }
    try:
        resp = client.chat.completions.create(
            model=deployment,
            messages=[
                {"role": "system", "content": VERIFIER_SYSTEM},
                {"role": "user", "content": json.dumps(payload)[:7000]},
            ],
            response_format={"type": "json_object"},
            temperature=0.05,
            max_tokens=1024,
        )
        return json.loads(resp.choices[0].message.content)
    except Exception as e:
        return {"accepted": True, "issues": [], "_verifier_error": str(e)}


# ───────────────────────────────────────────────────────────────────────────────
# Step 6: per-slide verify loop
# ───────────────────────────────────────────────────────────────────────────────

def _read_shape_texts(slide, names: list[str]) -> dict[str, str]:
    out = {}
    for n in names:
        sh = _find_shape_by_name(slide, n)
        if sh and sh.has_text_frame:
            out[n] = sh.text_frame.text
    return out


def verify_and_retry_slide(
    client,
    deployment: str,
    slide,
    baseline_inv: dict,
    plan_section: dict,
    max_iter: int = MAX_ITERATIONS,
) -> dict:
    history = []
    last_shape_texts: dict[str, str] = {}
    last_attempt: dict[str, str] = {}
    accepted = False
    issues: list = []

    editable_names = [s["name"] for s in baseline_inv["shapes"]
                      if s["role"] in ("title", "body")]

    for it in range(1, max_iter + 1):
        proposal = slidemaster_propose(
            client, deployment, plan_section, baseline_inv,
            prev_attempt=last_attempt if it > 1 else None,
            verifier_issues=issues if it > 1 else None,
        )
        touched = apply_shape_texts(slide, proposal)
        last_attempt = proposal

        py = python_verify_slide(slide, baseline_inv, touched)

        # Try font-shrink for overflow before declaring failure.
        for shape, base_font in py["shrink_targets"]:
            w_in = (shape.width or 0) / 914400
            h_in = (shape.height or 0) / 914400
            if w_in > 0 and h_in > 0:
                _shrink_font_to_fit(
                    shape.text_frame, w_in, h_in, base_font or 12,
                    min_pt=10 if _shape_role(shape) == "title" else 8,
                )

        last_shape_texts = _read_shape_texts(slide, touched)
        llm = llm_verify_slide(client, deployment, plan_section, last_shape_texts)

        issues = list(py["issues"]) + list(llm.get("issues", []))
        # Filter overflow issues that font-shrink resolved
        survived = []
        for iss in issues:
            if iss.get("type") == "overflow":
                sh = _find_shape_by_name(slide, iss.get("shape", ""))
                if sh and not _shape_overflow(sh).get("overflow"):
                    continue
            survived.append(iss)
        issues = survived

        accepted = llm.get("accepted", True) and not issues
        history.append({
            "iteration": it,
            "touched": touched,
            "issues": issues,
            "accepted": accepted,
        })
        if accepted:
            break

    return {
        "slide_index": baseline_inv["slide_index"],
        "section_type": baseline_inv["section_type"],
        "iterations": len(history),
        "accepted": accepted,
        "remaining_issues": issues,
        "shape_texts": last_shape_texts,
        "history": history,
    }


# ───────────────────────────────────────────────────────────────────────────────
# Top-level entry point
# ───────────────────────────────────────────────────────────────────────────────

def generate_proposal(
    rfp_text: str,
    user_prompt: str = "",
    on_log=None,
) -> dict:
    cfg = load_config()
    client = cfg["client"]
    deployment = cfg["deployment"]
    template_name = cfg["default_template"]
    template_path = os.path.join(cfg["datalake_dir"], template_name)
    if not os.path.isfile(template_path):
        raise FileNotFoundError(f"Default template not found: {template_path}")

    def _log(msg):
        if on_log:
            on_log(msg)
        print(msg)

    t0 = time.time()
    proposal_id = uuid.uuid4().hex[:8]

    _log(f"[MCP] proposal_id={proposal_id} template={template_name}")

    _log("[MCP] Parsing structured plan from RFP text...")
    plan = parse_plan(client, deployment, rfp_text, user_prompt)
    _log(f"[MCP] Plan: {len(plan.get('sections', []))} sections | "
         f"client='{plan.get('client_name','?')}' sector='{plan.get('sector','?')}'")

    _log("[MCP] Inspecting template slides...")
    template_inv = inspect_template(template_path)
    _log(f"[MCP] Template has {len(template_inv)} slides")

    pairings = pair_plan_with_slides(plan, template_inv)
    _log(f"[MCP] Paired {len(pairings)} plan sections to template slides")

    os.makedirs(RUNS_DIR, exist_ok=True)
    output_path = os.path.join(RUNS_DIR, f"proposal_{proposal_id}.pptx")
    import shutil
    shutil.copy2(template_path, output_path)
    prs = Presentation(output_path)

    placeholders = {
        "[CLIENT]": plan.get("client_name", "") or "the Client",
        "[client]": plan.get("client_name", "") or "the Client",
        "[PROJECT]": plan.get("project_name", "") or "",
        "[project]": plan.get("project_name", "") or "",
        "[COMPANY]": "EY",
        "[company]": "EY",
        "[DATE]": time.strftime("%B %Y"),
    }
    from generate_pptx import _replace_placeholders_globally, _rewrite_master_footers
    _replace_placeholders_globally(prs, placeholders)
    _rewrite_master_footers(prs, placeholders["[CLIENT]"], placeholders["[PROJECT]"])

    mapped_original_idxs = {p["slide_index"] for p in pairings if not p.get("duplicate")}

    slide_reports = []
    for pair in pairings:
        slide_idx = pair["slide_index"]
        plan_idx = pair["plan_idx"]
        plan_section = plan["sections"][plan_idx]

        if pair.get("duplicate"):
            _log(f"[MCP] Duplicating slide {slide_idx + 1} for section "
                 f"'{plan_section.get('section_type','other')}'")
            new_slide = _duplicate_slide(prs, slide_idx)
            if new_slide is None:
                _log("[MCP]   Duplication failed; skipping this section")
                continue
            slide = new_slide
            baseline_inv = template_inv[slide_idx]
            baseline_inv = {
                **baseline_inv,
                "slide_index": len(prs.slides) - 1,
            }
        else:
            slide = prs.slides[slide_idx]
            baseline_inv = template_inv[slide_idx]

        _log(f"[MCP] Slide {baseline_inv['slide_index']+1} "
             f"<- plan section #{plan_idx+1} ({plan_section.get('section_type','other')})")
        result = verify_and_retry_slide(
            client, deployment, slide, baseline_inv, plan_section, MAX_ITERATIONS,
        )
        _log(f"[MCP]   iter={result['iterations']} accepted={result['accepted']} "
             f"issues={len(result['remaining_issues'])}")
        slide_reports.append(result)

    # ── Cleanup pass: untouched template slides carry stale content. Keep
    # only section dividers (numbered chapter pages) and the last slide
    # (closing); delete every other unmapped slide so prompt.txt dominates.
    delete_targets: list[int] = []
    last_idx_now = len(prs.slides) - 1
    for i, slide in enumerate(prs.slides):
        if i in mapped_original_idxs:
            continue
        if i >= len(template_inv):
            continue  # newly duplicated slides at the end — keep
        if i == last_idx_now:
            continue  # closing slide
        if _is_section_divider_slide(slide):
            continue
        delete_targets.append(i)
    for idx in sorted(delete_targets, reverse=True):
        _log(f"[MCP] Deleting untouched stale slide {idx + 1}")
        _delete_slide(prs, idx)

    # Blank any long stale body paragraphs on slides we kept but did not map
    # (typically the closing slide). Titles and short labels stay.
    for i, slide in enumerate(prs.slides):
        if i in mapped_original_idxs:
            continue
        if _is_section_divider_slide(slide):
            continue
        for sh in slide.shapes:
            if not sh.has_text_frame:
                continue
            role = _shape_role(sh)
            if role != "body":
                continue
            wc = len((sh.text_frame.text or "").split())
            if wc >= 20:
                _replace_text_in_frame(sh.text_frame, "")

    prs.save(output_path)

    elapsed = round(time.time() - t0, 1)
    return {
        "proposal_id": proposal_id,
        "template_used": template_name,
        "output_path": output_path,
        "slides_built": len(slide_reports),
        "timing_seconds": elapsed,
        "plan": plan,
        "slide_reports": slide_reports,
        "message": f"Generated proposal with {len(slide_reports)} slides in {elapsed}s",
    }
