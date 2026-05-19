# Graph Report - C:\Users\MP946KB\WORKDIR\Proposal-assist\ProposalAssistant  (2026-05-15)

## Corpus Check
- Corpus is ~20,946 words - fits in a single context window. You may not need a graph.

## Summary
- 152 nodes · 207 edges · 13 communities (12 shown, 1 thin omitted)
- Extraction: 88% EXTRACTED · 11% INFERRED · 1% AMBIGUOUS · INFERRED: 23 edges (avg confidence: 0.81)
- Token cost: 0 input · 0 output

## Community Hubs (Navigation)
- [[_COMMUNITY_Knowledge Base Retrieval|Knowledge Base Retrieval]]
- [[_COMMUNITY_PPTX Slide Rendering|PPTX Slide Rendering]]
- [[_COMMUNITY_Slide Layout Verification|Slide Layout Verification]]
- [[_COMMUNITY_Web API and Run Logs|Web API and Run Logs]]
- [[_COMMUNITY_Config and Template Extraction|Config and Template Extraction]]
- [[_COMMUNITY_Flask Route Handlers|Flask Route Handlers]]
- [[_COMMUNITY_Section Mapping Pipeline|Section Mapping Pipeline]]
- [[_COMMUNITY_Mapping Debug Artifacts|Mapping Debug Artifacts]]
- [[_COMMUNITY_RFP Analysis Pipeline|RFP Analysis Pipeline]]
- [[_COMMUNITY_EY SVG Brand Assets|EY SVG Brand Assets]]
- [[_COMMUNITY_EY PNG Brand Assets|EY PNG Brand Assets]]
- [[_COMMUNITY_Package Initialization|Package Initialization]]

## God Nodes (most connected - your core abstractions)
1. `run_pipeline()` - 14 edges
2. `KnowledgeBase` - 10 edges
3. `main()` - 8 edges
4. `generate_pptx_from_enrichments()` - 6 edges
5. `_apply_content_to_slide()` - 6 edges
6. `classify_section()` - 6 edges
7. `verify_and_fix()` - 6 edges
8. `_fix_title_overflow()` - 6 edges
9. `_fix_body_overflow()` - 6 edges
10. `_replace_text_in_frame()` - 6 edges

## Surprising Connections (you probably didn't know these)
- `generate()` --calls--> `run_pipeline()`  [INFERRED]
  app.py → src/pipeline.py
- `RFP for Onboarding Digital Agency` --shares_data_with--> `RFP extraction and summarization`  [INFERRED]
  C:/Users/MP946KB/WORKDIR/Proposal-assist/ProposalAssistant/sampleinput/RFP for Onboarding Digital Agency.pdf → C:/Users/MP946KB/WORKDIR/Proposal-assist/ProposalAssistant/src/extract_rfp.py
- `Python dependency manifest` --references--> `Flask API entrypoint`  [INFERRED]
  C:/Users/MP946KB/WORKDIR/Proposal-assist/ProposalAssistant/requirements.txt → C:/Users/MP946KB/WORKDIR/Proposal-assist/ProposalAssistant/app.py
- `Debug mapping artifact` --references--> `Sample PPP RFP document`  [INFERRED]
  C:/Users/MP946KB/WORKDIR/Proposal-assist/ProposalAssistant/output/debug_mapping.json → C:/Users/MP946KB/WORKDIR/Proposal-assist/ProposalAssistant/sampleinput/sampleRFP 1.pdf
- `Pipeline run log #1` --references--> `Sample PPP RFP document`  [EXTRACTED]
  C:/Users/MP946KB/WORKDIR/Proposal-assist/ProposalAssistant/output/run_log.txt → C:/Users/MP946KB/WORKDIR/Proposal-assist/ProposalAssistant/sampleinput/sampleRFP 1.pdf

## Hyperedges (group relationships)
- **h_web_generation_flow** — n_static_index_html, n_app_py, n_src_pipeline_py, n_src_generate_pptx_py [EXTRACTED 1.00]
- **h_cli_mapping_flow** — n_src_main_py, n_src_extract_template_py, n_src_map_sections_py, n_output_debug_mapping_json [EXTRACTED 1.00]
- **h_sample_rfp_trace** — n_samplerfp1_pdf, n_output_run_log_txt, n_output_run_log2_txt, n_output_run_log3_txt, n_output_run_log4_txt [EXTRACTED 1.00]

## Communities (13 total, 1 thin omitted)

### Community 0 - "Knowledge Base Retrieval"
Cohesion: 0.11
Nodes (15): EmbeddingFunction, classify_section(), detect_sector(), KnowledgeBase, Knowledge Base: indexes all PPTX templates into ChromaDB for retrieval.  Hierarc, Parse and index all PPTX files in the datalake., Lightweight local embedding — no model download needed., Query the knowledge base with optional sector/section filtering. (+7 more)

### Community 1 - "PPTX Slide Rendering"
Cohesion: 0.14
Nodes (22): _apply_content_to_slide(), _duplicate_slide(), _find_shape_by_name(), _find_title_shape(), generate_pptx_from_enrichments(), _is_skip_shape(), Replace placeholders in a text frame, handling cross-run splits.          PPTX X, Apply enriched content to a slide with robust shape matching. (+14 more)

### Community 2 - "Slide Layout Verification"
Cohesion: 0.16
Nodes (21): _curate_slides(), _delete_slide(), _enforce_layout(), _estimate_max_chars_per_line(), _estimate_max_lines(), _fix_body_overflow(), _fix_title_overflow(), _get_dominant_font_size() (+13 more)

### Community 3 - "Web API and Run Logs"
Cohesion: 0.14
Nodes (20): Flask API entrypoint, Debug mapping artifact, Pipeline run log #2 (with traceback), Pipeline run log #3, Pipeline run log #4, Pipeline run log #1, Proposal generation prompt specification, Python dependency manifest (+12 more)

### Community 4 - "Config and Template Extraction"
Cohesion: 0.2
Nodes (10): load_config(), extract_rfp_text(), _classify_slides_with_llm(), _collect_slide_info(), extract_template_structure(), list_files(), main(), pick_option() (+2 more)

### Community 5 - "Flask Route Handlers"
Cohesion: 0.17
Nodes (9): customize(), download(), generate(), list_templates(), Flask web application for Proposal Assist., Download the generated PPTX file., Re-run pipeline with customization instructions (placeholder for POC)., Return available base templates from datalake/. (+1 more)

### Community 6 - "Section Mapping Pipeline"
Cohesion: 0.27
Nodes (10): _build_slide_outline_from_list(), _compute_qa(), _extract_template_slides(), _extract_template_slides_from_prs(), _format_rfp_summary(), _get_kb(), Agentic pipeline with Knowledge Base retrieval and RFP summary., Extract slide info from an already-open Presentation object. (+2 more)

### Community 7 - "Mapping Debug Artifacts"
Cohesion: 0.22
Nodes (8): adapted_keys, mapping, mappings, new_slides_needed, notes, slides_to_keep_original, rfp_sections, template_slides

### Community 8 - "RFP Analysis Pipeline"
Cohesion: 0.53
Nodes (5): analyze_rfp(), _build_knowledge_base(), _chunk_text(), _create_rfp_summary(), Extract key facts and create a knowledge base from the RFP.

### Community 9 - "EY SVG Brand Assets"
Cohesion: 0.5
Nodes (4): Left white mark, EY logo (SVG), Right white mark, Yellow beam

### Community 10 - "EY PNG Brand Assets"
Cohesion: 0.5
Nodes (3): Dark charcoal background, EY white wordmark, Yellow triangular wedge above wordmark

## Ambiguous Edges - Review These
- `RFP extraction and summarization` → `CLI orchestrator pipeline`  [AMBIGUOUS]
  C:/Users/MP946KB/WORKDIR/Proposal-assist/ProposalAssistant/src/main.py · relation: calls
- `PPTX generation and text replacement` → `CLI orchestrator pipeline`  [AMBIGUOUS]
  C:/Users/MP946KB/WORKDIR/Proposal-assist/ProposalAssistant/src/main.py · relation: calls

## Knowledge Gaps
- **59 isolated node(s):** `Flask web application for Proposal Assist.`, `Return available base templates from datalake/.`, `Accept RFP + template choice, run pipeline, return result.`, `Download the generated PPTX file.`, `Re-run pipeline with customization instructions (placeholder for POC).` (+54 more)
  These have ≤1 connection - possible missing edges or undocumented components.
- **1 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **What is the exact relationship between `RFP extraction and summarization` and `CLI orchestrator pipeline`?**
  _Edge tagged AMBIGUOUS (relation: calls) - confidence is low._
- **What is the exact relationship between `PPTX generation and text replacement` and `CLI orchestrator pipeline`?**
  _Edge tagged AMBIGUOUS (relation: calls) - confidence is low._
- **Why does `run_pipeline()` connect `Section Mapping Pipeline` to `Knowledge Base Retrieval`, `PPTX Slide Rendering`, `Slide Layout Verification`, `Config and Template Extraction`, `Flask Route Handlers`, `RFP Analysis Pipeline`?**
  _High betweenness centrality (0.406) - this node is a cross-community bridge._
- **Why does `generate_pptx_from_enrichments()` connect `PPTX Slide Rendering` to `Section Mapping Pipeline`?**
  _High betweenness centrality (0.177) - this node is a cross-community bridge._
- **Why does `extract_rfp_text()` connect `Config and Template Extraction` to `RFP Analysis Pipeline`, `Section Mapping Pipeline`?**
  _High betweenness centrality (0.120) - this node is a cross-community bridge._
- **Are the 6 inferred relationships involving `run_pipeline()` (e.g. with `generate()` and `extract_rfp_text()`) actually correct?**
  _`run_pipeline()` has 6 INFERRED edges - model-reasoned connections that need verification._
- **Are the 5 inferred relationships involving `main()` (e.g. with `load_config()` and `extract_rfp_text()`) actually correct?**
  _`main()` has 5 INFERRED edges - model-reasoned connections that need verification._