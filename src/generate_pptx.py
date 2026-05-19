import copy
import os
import re
import shutil
from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

_CTRL_CHAR_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")
SKIP_PATTERNS = {"slide number", "date placeholder", "footer placeholder",
                 "freeform", "connector", "straight connector"}


def _rewrite_master_footers(prs, client_name: str, project_name: str):
    """Overwrite master/layout footer shapes that still carry a stale project tagline.

    Master footers like `Footer Placeholder 2` in the template typically read
    `[CLIENT] : Healthcare Planning Consultant`. After placeholder replacement the
    `[CLIENT]` is fixed, but the hardcoded project descriptor remains and points at
    the wrong engagement. We rewrite any master/layout footer whose name starts with
    `Footer Placeholder` to `{client_name} : {project_name}` when we have both.
    """
    if not client_name and not project_name:
        return
    new_text_full = f"{client_name} : {project_name}".strip(" :") if (client_name and project_name) else (client_name or project_name)
    rewrote = 0
    targets = []
    for master in prs.slide_masters:
        targets.append(master.shapes)
        for lay in master.slide_layouts:
            targets.append(lay.shapes)
    for shapes in targets:
        for sh in shapes:
            if not sh.has_text_frame:
                continue
            name_l = sh.name.lower()
            if not name_l.startswith("footer placeholder"):
                continue
            cur = sh.text_frame.text or ""
            if not cur.strip() or cur.strip() in ("a", "<#>", "‹#›"):
                continue
            if cur.strip() == new_text_full:
                continue
            _replace_text_preserving_format(sh.text_frame, new_text_full)
            rewrote += 1
    if rewrote:
        print(f"    Rewrote {rewrote} master/layout footer(s) -> '{new_text_full[:80]}'")


def generate_pptx_from_enrichments(
    template_path: str,
    output_dir: str,
    content_map: dict,
    placeholders: dict,
) -> str:
    """
    Clone the template and apply:
    1. Global placeholder replacement ([CLIENT], [COMPANY], etc.)
    2. Content enrichments from the RFP for specific slides
    3. Slide duplication for overflow content

    content_map: {slide_index(int): {"Shape Name": "new text", ...}}
        The agent returns per-shape content. Each shape either gets new text or is skipped.
    placeholders: {"[CLIENT]": "Actual Name", ...}
    """
    base_name = os.path.splitext(os.path.basename(template_path))[0]
    output_path = os.path.join(output_dir, f"{base_name} - RFP Response.pptx")
    os.makedirs(output_dir, exist_ok=True)
    shutil.copy2(template_path, output_path)

    prs = Presentation(output_path)

    # ── Phase 1: Apply per-shape enrichments ──
    for slide_idx, shapes_content in content_map.items():
        slide_idx = int(slide_idx)
        if slide_idx >= len(prs.slides) or not shapes_content:
            continue

        slide = prs.slides[slide_idx]
        applied = 0
        for shape_name, new_text in shapes_content.items():
            if not new_text or not isinstance(new_text, str):
                continue
            if new_text.strip().upper() == "KEEP":
                continue
            shape = _find_shape_by_name(slide, shape_name)
            if shape and shape.has_text_frame:
                _replace_text_preserving_format(shape.text_frame, new_text)
                applied += 1

        if applied > 0:
            print(f"    Slide {slide_idx + 1}: updated {applied} shapes")

    # ── Phase 2: Global placeholder replacement AFTER all content ──
    ph_debug = {k: v[:30] for k, v in placeholders.items() if v}
    print(f"    Replacing placeholders: {ph_debug}")
    _replace_placeholders_globally(prs, placeholders)

    # ── Phase 3: Rewrite stale master/layout footer taglines ──
    client_name = placeholders.get("[CLIENT]") or placeholders.get("[client]") or ""
    project_name = placeholders.get("[PROJECT]") or placeholders.get("[project]") or ""
    _rewrite_master_footers(prs, client_name, project_name)

    prs.save(output_path)
    return output_path


# ---------------------------------------------------------------------------
# Global placeholder replacement
# ---------------------------------------------------------------------------


def _replace_placeholders_globally(prs, placeholders: dict):
    """Replace [CLIENT], [COMPANY], etc. in ALL text across slides, layouts, and masters.

    Footers and confidential tags often live on slide masters / layouts, not on slides
    themselves, so we must walk all three to catch every occurrence.
    """
    if not placeholders:
        return

    replaced = 0
    for slide in prs.slides:
        replaced += _replace_in_shapes(slide.shapes, placeholders)

    seen_layouts = set()
    for slide in prs.slides:
        lay = slide.slide_layout
        if id(lay) in seen_layouts:
            continue
        seen_layouts.add(id(lay))
        replaced += _replace_in_shapes(lay.shapes, placeholders)

    for master in prs.slide_masters:
        replaced += _replace_in_shapes(master.shapes, placeholders)
        for lay in master.slide_layouts:
            if id(lay) in seen_layouts:
                continue
            seen_layouts.add(id(lay))
            replaced += _replace_in_shapes(lay.shapes, placeholders)

    print(f"    Replaced placeholders in {replaced} text frames (slides + layouts + masters)")


def _replace_in_shapes(shapes, placeholders: dict) -> int:
    """Recursively replace placeholders in shapes, including group children."""
    count = 0
    for shape in shapes:
        if shape.shape_type == 6:  # MSO_SHAPE_TYPE.GROUP
            try:
                count += _replace_in_shapes(shape.shapes, placeholders)
            except Exception:
                pass
        if shape.has_text_frame:
            if _replace_placeholders_in_textframe(shape.text_frame, placeholders):
                count += 1
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    if _replace_placeholders_in_textframe(cell.text_frame, placeholders):
                        count += 1
    return count


def _replace_placeholders_in_textframe(text_frame, placeholders: dict) -> bool:
    """Replace placeholders in a text frame, handling cross-run splits.
    
    PPTX XML can split [CLIENT] across runs like: [  |  CLIENT  |  ]
    So we must join all run texts, do the replacement, then redistribute.
    Returns True if any replacement was made.
    """
    changed = False
    for paragraph in text_frame.paragraphs:
        runs = paragraph.runs
        if not runs:
            continue

        # Join all run texts
        full_text = "".join(r.text or "" for r in runs)
        if not full_text:
            continue

        new_text = full_text
        # Case-insensitive placeholder matching
        for placeholder, replacement in placeholders.items():
            if placeholder.lower() in new_text.lower():
                # Replace all case variants
                idx = 0
                result = ""
                text_lower = new_text.lower()
                ph_lower = placeholder.lower()
                while idx < len(new_text):
                    pos = text_lower.find(ph_lower, idx)
                    if pos == -1:
                        result += new_text[idx:]
                        break
                    result += new_text[idx:pos] + replacement
                    idx = pos + len(placeholder)
                new_text = result

        if new_text == full_text:
            continue

        changed = True
        # Text changed -> put all text in first run, clear the rest
        first_t = runs[0]._r.find(f"{{{NS_A}}}t")
        if first_t is not None:
            first_t.text = _sanitize(new_text)

        for run in runs[1:]:
            t_elem = run._r.find(f"{{{NS_A}}}t")
            if t_elem is not None:
                t_elem.text = ""

    return changed


# ---------------------------------------------------------------------------
# Content application
# ---------------------------------------------------------------------------


def _apply_content_to_slide(slide, page_content: dict):
    """Apply enriched content to a slide with robust shape matching."""
    applied_shapes = set()

    # Apply title
    title_text = page_content.get("title")
    if title_text:
        title_shape = _find_title_shape(slide)
        if title_shape and title_shape.has_text_frame:
            _replace_text_preserving_format(title_shape.text_frame, title_text)
            applied_shapes.add(title_shape.name)

    # Apply body parts by exact name
    body_parts = page_content.get("body_parts", [])
    unmatched_parts = []

    for part in body_parts:
        shape_name = part.get("shape_name")
        content = part.get("content", "")
        if not content:
            continue

        shape = _find_shape_by_name(slide, shape_name) if shape_name else None
        if shape and shape.has_text_frame and shape.name not in applied_shapes:
            if not _is_skip_shape(shape.name):
                _replace_text_preserving_format(shape.text_frame, content)
                applied_shapes.add(shape.name)
                continue

        unmatched_parts.append(content)

    # Fallback: apply unmatched content to largest unused text shapes
    if unmatched_parts:
        candidates = []
        for shape in slide.shapes:
            if not shape.has_text_frame or shape.name in applied_shapes:
                continue
            if _is_skip_shape(shape.name):
                continue
            wc = len(shape.text_frame.text.split())
            if wc >= 3:
                candidates.append((wc, shape))
        candidates.sort(key=lambda x: x[0], reverse=True)

        for content, (_, shape) in zip(unmatched_parts, candidates):
            _replace_text_preserving_format(shape.text_frame, content)
            applied_shapes.add(shape.name)


def _find_title_shape(slide):
    """Find the title shape on a slide."""
    # Priority 1: shape with "title" in name
    for shape in slide.shapes:
        if shape.has_text_frame and "title" in shape.name.lower():
            if not _is_skip_shape(shape.name):
                return shape
    # Priority 2: shape.shapes.title
    if slide.shapes.title and slide.shapes.title.has_text_frame:
        return slide.shapes.title
    return None


def _find_shape_by_name(slide, name: str):
    """Find shape by name, handling disambiguated names like 'Content Placeholder 8 #3'."""
    if not name:
        return None
    # Check if name has a disambiguation suffix
    dup_match = re.match(r'^(.+?) #(\d+)$', name)
    if dup_match:
        base_name = dup_match.group(1)
        occurrence = int(dup_match.group(2))
        count = 0
        for shape in slide.shapes:
            if shape.name == base_name:
                count += 1
                if count == occurrence:
                    return shape
        return None
    # Direct name match
    for shape in slide.shapes:
        if shape.name == name:
            return shape
    return None


def _is_skip_shape(name: str) -> bool:
    name_lower = name.lower()
    return any(s in name_lower for s in SKIP_PATTERNS)


# ---------------------------------------------------------------------------
# XML-level text replacement (preserves ALL formatting)
# ---------------------------------------------------------------------------


def _replace_text_preserving_format(text_frame, new_text: str):
    """Replace text while preserving fonts, colors, sizes, bullets."""
    txBody = text_frame._txBody
    a_p = f"{{{NS_A}}}p"
    a_r = f"{{{NS_A}}}r"
    a_t = f"{{{NS_A}}}t"
    a_rPr = f"{{{NS_A}}}rPr"
    a_pPr = f"{{{NS_A}}}pPr"
    a_endParaRPr = f"{{{NS_A}}}endParaRPr"

    existing_paras = [child for child in txBody if child.tag == a_p]
    if not existing_paras:
        return

    # Detect: did the shape have any real text BEFORE this call? If empty, the
    # paragraph properties we inherit may come from layout defaults (often
    # CENTER on autoshapes) which are wrong for body text. We force-rewrite
    # alignment to LEFT for body-content paragraphs in that case.
    had_text_before = any(
        ("".join((r.find(a_t).text or "") for r in p.findall(a_r) if r.find(a_t) is not None)).strip()
        for p in existing_paras
    )

    # Collect formatting from existing paragraphs
    para_formats = []
    for p in existing_paras:
        pPr = p.find(a_pPr)
        runs = p.findall(a_r)
        endParaRPr = p.find(a_endParaRPr)
        rPr = None
        if runs:
            rPr = runs[0].find(a_rPr)
        para_formats.append({
            "pPr": copy.deepcopy(pPr) if pPr is not None else None,
            "rPr": copy.deepcopy(rPr) if rPr is not None else None,
            "endParaRPr": copy.deepcopy(endParaRPr) if endParaRPr is not None else None,
        })

    # If the shape was empty, body-content paragraphs should default to LEFT
    # (most autoshape bodies inherit CENTER from layout, which looks wrong for
    # prose). We only override alignment, leaving font/colour untouched.
    if not had_text_before:
        for fmt in para_formats:
            if fmt["pPr"] is None:
                fmt["pPr"] = etree.Element(a_pPr)
            fmt["pPr"].set("algn", "l")

    # Body format: use 2nd paragraph if available (first is often a sub-heading)
    body_fmt_idx = min(1, len(para_formats) - 1)
    body_format = para_formats[body_fmt_idx]

    # Split new text into lines
    new_lines = [line for line in new_text.split("\n") if line.strip() or line == ""]
    if not new_lines:
        new_lines = [""]

    # Remove existing paragraphs
    for p in existing_paras:
        txBody.remove(p)

    # Build new paragraphs with preserved formatting
    for i, line_text in enumerate(new_lines):
        fmt = para_formats[i] if i < len(para_formats) else body_format

        p_elem = etree.SubElement(txBody, a_p)
        if fmt["pPr"] is not None:
            p_elem.insert(0, copy.deepcopy(fmt["pPr"]))

        r_elem = etree.SubElement(p_elem, a_r)
        if fmt["rPr"] is not None:
            r_elem.insert(0, copy.deepcopy(fmt["rPr"]))

        t_elem = etree.SubElement(r_elem, a_t)
        t_elem.text = _sanitize(line_text)
        t_elem.set(XML_SPACE, "preserve")

        if fmt["endParaRPr"] is not None:
            p_elem.append(copy.deepcopy(fmt["endParaRPr"]))


# ---------------------------------------------------------------------------
# Slide duplication
# ---------------------------------------------------------------------------


def _duplicate_slide(prs, source_idx: int):
    """Duplicate a slide preserving all shapes, images, and relationships."""
    if source_idx >= len(prs.slides):
        return None

    source = prs.slides[source_idx]
    try:
        new_slide = prs.slides.add_slide(source.slide_layout)
    except Exception as e:
        print(f"    Warning: Could not add slide: {e}")
        return None

    src_elem = source._element
    new_elem = new_slide._element

    # Build rId mapping
    rId_map = {}
    for new_rId_key, new_rel in new_slide.part.rels.items():
        if new_rel.reltype == RT.SLIDE_LAYOUT:
            for src_rId_key, src_rel in source.part.rels.items():
                if src_rel.reltype == RT.SLIDE_LAYOUT:
                    rId_map[src_rId_key] = new_rId_key
                    break
            break

    for src_rId, src_rel in source.part.rels.items():
        if src_rel.reltype == RT.SLIDE_LAYOUT:
            continue
        try:
            if src_rel.is_external:
                new_rId = new_slide.part.relate_to(
                    src_rel.target_ref, src_rel.reltype, is_external=True
                )
            else:
                new_rId = new_slide.part.relate_to(
                    src_rel.target_part, src_rel.reltype
                )
            rId_map[src_rId] = new_rId
        except Exception:
            rId_map[src_rId] = src_rId

    # Copy slide content
    src_cSld = src_elem.find(f"{{{NS_P}}}cSld")
    new_cSld = new_elem.find(f"{{{NS_P}}}cSld")

    if src_cSld is not None and new_cSld is not None:
        parent = new_cSld.getparent()
        idx = list(parent).index(new_cSld)
        parent.remove(new_cSld)
        cloned = copy.deepcopy(src_cSld)
        _remap_rids(cloned, rId_map)
        parent.insert(idx, cloned)

    return new_slide


def _remap_rids(element, rId_map: dict):
    r_id_attr = f"{{{NS_R}}}id"
    r_embed_attr = f"{{{NS_R}}}embed"
    r_link_attr = f"{{{NS_R}}}link"
    for elem in element.iter():
        for attr in (r_id_attr, r_embed_attr, r_link_attr):
            old_val = elem.get(attr)
            if old_val and old_val in rId_map:
                elem.set(attr, rId_map[old_val])


def _sanitize(text: str) -> str:
    """Remove control characters that are invalid in XML."""
    return _CTRL_CHAR_RE.sub("", text) if text else ""
