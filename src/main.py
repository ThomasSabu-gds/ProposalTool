import os
import sys
import json
import time

from config import load_config
from extract_rfp import extract_rfp_text, parse_rfp_sections
from extract_template import extract_template_structure
from map_sections import map_sections, adapt_content_for_slides
from generate_pptx import generate_pptx


def list_files(directory: str, extensions: tuple) -> list[str]:
    files = []
    for f in sorted(os.listdir(directory)):
        if f.lower().endswith(extensions):
            files.append(f)
    return files


def pick_option(prompt: str, options: list[str]) -> int:
    print(f"\n{prompt}")
    for i, opt in enumerate(options, 1):
        print(f"  {i}. {opt}")
    while True:
        try:
            choice = int(input("\nEnter choice number: "))
            if 1 <= choice <= len(options):
                return choice - 1
        except (ValueError, EOFError):
            pass
        print(f"  Please enter a number between 1 and {len(options)}")


def main():
    print("=" * 60)
    print("  RFP-to-Presentation Generator")
    print("=" * 60)

    config = load_config()
    client = config["client"]
    deployment = config["deployment"]

    # Step 1: Choose template
    templates = list_files(config["datalake_dir"], (".pptx",))
    if not templates:
        print("ERROR: No .pptx templates found in datalake/")
        sys.exit(1)
    template_idx = pick_option("Select a base template:", templates)
    template_path = os.path.join(config["datalake_dir"], templates[template_idx])
    print(f"  -> Using: {templates[template_idx]}")

    # Step 2: Choose RFP
    rfps = list_files(config["input_dir"], (".pdf",))
    if not rfps:
        print("ERROR: No .pdf RFP files found in sampleinput/")
        sys.exit(1)
    rfp_idx = pick_option("Select an RFP document:", rfps)
    rfp_path = os.path.join(config["input_dir"], rfps[rfp_idx])
    print(f"  -> Using: {rfps[rfp_idx]}")

    t0 = time.time()

    # Step 3: Extract RFP
    print("\n[1/5] Extracting RFP content...")
    rfp_text = extract_rfp_text(rfp_path)
    rfp_sections = parse_rfp_sections(client, deployment, rfp_text)
    print(f"  Found {len(rfp_sections)} sections in RFP:")
    for sec in rfp_sections[:10]:
        sub_count = len(sec.get("subsections", []))
        print(f"    - {sec['heading']}" + (f" ({sub_count} subsections)" if sub_count else ""))
    if len(rfp_sections) > 10:
        print(f"    ... and {len(rfp_sections) - 10} more")

    # Step 4: Analyze template
    print("\n[2/5] Analyzing template structure...")
    template_structure = extract_template_structure(template_path, client, deployment)
    content_slides = [
        s for s in template_structure
        if s.get("slide_purpose") == "content"
        and any(cs.get("replaceable") for cs in s.get("content_shapes", []))
    ]
    print(f"  Analyzed {len(template_structure)} slides, {len(content_slides)} are replaceable content slides")

    # Step 5: Map sections
    print("\n[3/5] Mapping RFP sections to template slides...")
    mapping = map_sections(client, deployment, template_structure, rfp_sections)
    mappings_list = mapping.get("mappings", [])
    print(f"  Created {len(mappings_list)} section mappings")
    new_slides = mapping.get("new_slides_needed", [])
    if new_slides:
        print(f"  {len(new_slides)} new slides will be created for unmapped sections")

    for m in mappings_list:
        t_idx = m.get("template_slide_index", "?")
        r_idx = m.get("rfp_section_index", "?")
        heading = "?"
        if isinstance(r_idx, int) and r_idx < len(rfp_sections):
            heading = rfp_sections[r_idx]["heading"]
        print(f"    Slide {t_idx + 1 if isinstance(t_idx, int) else t_idx} <- {heading}")

    # Step 6: Adapt content
    print("\n[4/5] Preparing slide-ready content...")
    adapted_content = {}

    for m in mappings_list:
        t_idx = m.get("template_slide_index")
        r_idx = m.get("rfp_section_index")
        r_sub_idx = m.get("rfp_subsection_index")

        if t_idx is None or r_idx is None:
            continue
        if r_idx >= len(rfp_sections):
            continue

        rfp_sec = rfp_sections[r_idx]
        if r_sub_idx is not None:
            subs = rfp_sec.get("subsections", [])
            if r_sub_idx < len(subs):
                rfp_sec = subs[r_sub_idx]

        # Get body shapes for this template slide
        slide_info = None
        for s in template_structure:
            if s["index"] == t_idx:
                slide_info = s
                break
        if not slide_info:
            continue

        body_shapes = [
            cs for cs in slide_info.get("content_shapes", [])
            if cs.get("role") == "body" and cs.get("replaceable")
        ]
        if not body_shapes:
            # Fall back to any replaceable shape
            body_shapes = [
                cs for cs in slide_info.get("content_shapes", [])
                if cs.get("replaceable")
            ]
        if not body_shapes:
            continue

        print(f"    Adapting content for slide {t_idx + 1}: {rfp_sec['heading'][:50]}...")
        try:
            slides_content = adapt_content_for_slides(
                client, deployment, rfp_sec, slide_info, body_shapes
            )
            adapted_content[str(t_idx)] = slides_content
        except Exception as e:
            print(f"    Warning: Failed to adapt content for slide {t_idx + 1}: {e}")

    # Also adapt content for new slides
    for ns_info in new_slides:
        r_idx = ns_info.get("rfp_section_index")
        clone_from = ns_info.get("clone_from_slide_index", 0)

        if r_idx is None or r_idx >= len(rfp_sections):
            continue

        rfp_sec = rfp_sections[r_idx]
        slide_info = None
        for s in template_structure:
            if s["index"] == clone_from:
                slide_info = s
                break
        if not slide_info:
            continue

        body_shapes = [
            cs for cs in slide_info.get("content_shapes", [])
            if cs.get("replaceable")
        ]
        if not body_shapes:
            continue

        print(f"    Adapting new slide content: {rfp_sec['heading'][:50]}...")
        try:
            slides_content = adapt_content_for_slides(
                client, deployment, rfp_sec, slide_info, body_shapes
            )
            adapted_content[f"new_{r_idx}"] = slides_content
        except Exception as e:
            print(f"    Warning: Failed to adapt new slide content: {e}")

    # Step 7: Generate PPTX
    print("\n[5/5] Generating presentation...")
    output_path = generate_pptx(
        template_path=template_path,
        output_dir=config["output_dir"],
        template_structure=template_structure,
        mapping=mapping,
        rfp_sections=rfp_sections,
        adapted_content=adapted_content,
    )

    elapsed = time.time() - t0
    print(f"\n{'=' * 60}")
    print(f"  DONE in {elapsed:.1f}s")
    print(f"  Output: {output_path}")
    print(f"{'=' * 60}")

    # Save debug info
    debug_path = os.path.join(config["output_dir"], "debug_mapping.json")
    with open(debug_path, "w", encoding="utf-8") as f:
        json.dump(
            {
                "rfp_sections": [
                    {"heading": s["heading"], "subsections": [sub["heading"] for sub in s.get("subsections", [])]}
                    for s in rfp_sections
                ],
                "template_slides": [
                    {"index": s["index"], "heading": s.get("section_heading", ""), "purpose": s.get("slide_purpose", "")}
                    for s in template_structure
                ],
                "mapping": mapping,
                "adapted_slides": list(adapted_content.keys()),
            },
            f,
            indent=2,
            ensure_ascii=False,
        )
    print(f"  Debug info: {debug_path}")


if __name__ == "__main__":
    main()
