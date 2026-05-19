import json


MAP_SYSTEM = """You are a presentation content strategist. Your job is to FILL a template presentation with RFP content.

The template defines the VISUAL LAYOUT and STYLE. The RFP provides the NEW CONTENT.
Your goal: maximize the number of template slides populated with RFP content.

Return JSON:
{
  "mappings": [
    {
      "template_slide_index": 0,
      "rfp_section_index": 2,
      "rfp_subsection_index": null,
      "confidence": "high",
      "reasoning": "Brief reason"
    }
  ],
  "new_slides_needed": [
    {
      "rfp_section_index": 5,
      "rfp_subsection_index": null,
      "clone_from_slide_index": 3,
      "reasoning": "Important RFP section needs a new slide"
    }
  ],
  "slides_to_keep_original": [0, 1],
  "notes": "Summary"
}

MAPPING STRATEGY (in order of priority):
1. TOPIC MATCH: If an RFP section and template slide share a similar topic, map them.
2. STRUCTURAL FIT: Map RFP sections to template slides by position:
   - Early template content slides -> RFP intro, background, overview sections
   - Middle template content slides -> RFP core requirements, process, evaluation sections
   - Late template content slides -> RFP appendices, legal, miscellaneous sections
3. FILL REMAINING: Any replaceable content slide not yet mapped should get the next unmapped RFP section (in order).

Rules:
- ONLY keep slides as original if they are: title_slide, agenda, section_header, transition, closing, or appendix slides
- ALL content slides with has_replaceable_body=true SHOULD be mapped to an RFP section
- A template slide can only map to ONE RFP section (but a long RFP section can be split across multiple template slides)
- If RFP sections outnumber template slides, create new_slides_needed entries using the best layout
- Aim to produce at least 10-15 mappings for a typical template"""


def map_sections(
    client,
    deployment: str,
    template_structure: list[dict],
    rfp_sections: list[dict],
) -> dict:
    # Build compact summaries for the LLM
    template_summary = []
    for slide in template_structure:
        template_summary.append(
            {
                "index": slide["index"],
                "section_heading": slide.get("section_heading", ""),
                "slide_purpose": slide.get("slide_purpose", ""),
                "has_replaceable_body": any(
                    cs.get("replaceable")
                    for cs in slide.get("content_shapes", [])
                ),
                "replaceable_shapes": [
                    cs["name"] for cs in slide.get("content_shapes", [])
                    if cs.get("replaceable")
                ],
            }
        )

    rfp_summary = []
    for i, sec in enumerate(rfp_sections):
        content_preview = (sec.get("content") or "")[:200]
        subsections = []
        for j, sub in enumerate(sec.get("subsections", [])):
            subsections.append(
                {
                    "index": j,
                    "heading": sub["heading"],
                    "content_preview": (sub.get("content") or "")[:150],
                }
            )
        rfp_summary.append(
            {
                "index": i,
                "heading": sec["heading"],
                "content_preview": content_preview,
                "subsections": subsections,
            }
        )

    prompt = (
        f"TEMPLATE SLIDES:\n{json.dumps(template_summary, indent=2)}\n\n"
        f"RFP SECTIONS:\n{json.dumps(rfp_summary, indent=2)}"
    )

    resp = client.chat.completions.create(
        model=deployment,
        messages=[
            {"role": "system", "content": MAP_SYSTEM},
            {"role": "user", "content": prompt},
        ],
        response_format={"type": "json_object"},
        temperature=0.05,
        max_tokens=8000,
    )

    mapping = json.loads(resp.choices[0].message.content)
    return mapping


ADAPT_SYSTEM = """You are a presentation content writer. Transform RFP document text into polished slide content.

You receive:
- The original template slide content (showing the format and style to match)
- The RFP section content to transform
- Format hints (bullet style, word count target, paragraph count)

Return JSON:
{
  "slides": [
    {
      "title": "Slide heading text",
      "body_parts": [
        {
          "shape_name": "name of the shape to put this content into",
          "content": "The formatted slide content with line breaks as \\n"
        }
      ]
    }
  ]
}

Rules:
- Match the STYLE of the original: if it uses bullets, use bullets. If short phrases, use short phrases.
- Target approximately the same word count as the original (+/- 30%)
- If content is >1.5x the original word count, SPLIT into multiple slides
- Each slide must be self-contained and readable
- Use professional, concise language appropriate for executive presentations
- Preserve ALL key facts, numbers, dates, and requirements from the RFP
- Do NOT add placeholder text like [Company Name] unless the original had it
- First slide's title should match the RFP section heading
- Additional slides (for overflow) should have "... (continued)" in title"""


def adapt_content_for_slides(
    client,
    deployment: str,
    rfp_section: dict,
    template_slide: dict,
    body_shapes: list[dict],
    extra_system: str = "",
) -> list[dict]:
    # Build the original content reference
    original_content = {}
    for cs in body_shapes:
        original_content[cs["name"]] = {
            "text": cs.get("_original_text", "")[:500],
            "word_count": cs.get("_word_count", 50),
            "has_bullets": cs.get("_has_bullets", False),
            "para_count": cs.get("_para_count", 1),
        }

    # Get the RFP content (include subsections)
    rfp_content = rfp_section.get("content", "")
    for sub in rfp_section.get("subsections", []):
        rfp_content += f"\n\n{sub['heading']}:\n{sub.get('content', '')}"

    total_target_words = sum(v["word_count"] for v in original_content.values())
    total_target_words = max(total_target_words, 40)

    prompt = (
        f"ORIGINAL TEMPLATE CONTENT (match this style):\n"
        f"{json.dumps(original_content, indent=2, default=str)}\n\n"
        f"RFP SECTION: {rfp_section['heading']}\n"
        f"RFP CONTENT:\n{rfp_content[:6000]}\n\n"
        f"TARGET: ~{total_target_words} words per slide. "
        f"Shape names to fill: {[s['name'] for s in body_shapes]}"
    )

    system_prompt = ADAPT_SYSTEM
    if extra_system:
        system_prompt += extra_system

    resp = client.chat.completions.create(
        model=deployment,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": prompt},
        ],
        response_format={"type": "json_object"},
        temperature=0.15,
        max_tokens=8000,
    )

    result = json.loads(resp.choices[0].message.content)
    return result.get("slides", [])
