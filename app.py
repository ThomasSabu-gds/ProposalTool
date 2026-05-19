"""Flask web application for Proposal Assist."""

import os
import sys
import uuid
import traceback

from flask import Flask, request, jsonify, send_file, send_from_directory

# Add src to path so pipeline can import its siblings
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from config import load_config
from pipeline import run_pipeline
from template_preview import (
    get_preview_path,
    get_template_metadata,
    get_layout_previews,
    warm_layout_cache_async,
)

app = Flask(__name__, static_folder="static")
cfg = load_config()
STATIC_DIR = os.path.join(os.path.dirname(__file__), "static")

# In-memory proposal store (fine for POC)
proposals: dict = {}


# ─── Static routes ───────────────────────────────────────────────────────────

@app.route("/")
def index():
    return send_from_directory("static", "index.html")


@app.route("/static/<path:path>")
def serve_static(path):
    return send_from_directory("static", path)


# ─── API routes ──────────────────────────────────────────────────────────────

@app.route("/templates", methods=["GET"])
def list_templates():
    """Return available base templates from datalake/ with metadata for the gallery."""
    files = sorted(
        f for f in os.listdir(cfg["datalake_dir"]) if f.lower().endswith(".pptx")
    )
    items = []
    for fname in files:
        tpl_path = os.path.join(cfg["datalake_dir"], fname)
        meta = get_template_metadata(tpl_path)
        items.append(
            {
                "name": fname,
                "label": os.path.splitext(fname)[0],
                "preview_url": f"/templates/preview/{fname}",
                "slide_count": meta.get("slide_count", 0),
                "sector": meta.get("sector", "General"),
                "accent_color": meta.get("accent_color", "#FFE600"),
            }
        )
    return jsonify({"templates": items})


@app.route("/templates/preview/<path:filename>", methods=["GET"])
def template_preview(filename: str):
    """Serve a cached PNG preview of the template (generates on demand)."""
    safe_name = os.path.basename(filename)
    template_path = os.path.join(cfg["datalake_dir"], safe_name)
    if not os.path.isfile(template_path):
        return jsonify({"detail": "Template not found"}), 404
    png_path = get_preview_path(template_path, STATIC_DIR)
    if not png_path or not os.path.isfile(png_path):
        return jsonify({"detail": "Could not render preview"}), 500
    return send_file(png_path, mimetype="image/png", max_age=3600)


@app.route("/templates/layouts/<path:filename>", methods=["GET"])
def template_layouts(filename: str):
    """List empty master-layout previews for a template (gallery on click)."""
    safe_name = os.path.basename(filename)
    template_path = os.path.join(cfg["datalake_dir"], safe_name)
    if not os.path.isfile(template_path):
        return jsonify({"detail": "Template not found"}), 404
    layouts = get_layout_previews(template_path, STATIC_DIR)
    return jsonify(
        {
            "template": safe_name,
            "layouts": [
                {"index": l["index"], "name": l["name"], "preview_url": l["rel_url"]}
                for l in layouts
            ],
        }
    )


@app.route("/generate", methods=["POST"])
def generate():
    """Accept RFP + template choice, run pipeline, return result."""
    rfp_file = request.files.get("rfp_file")
    template_name = request.form.get("template_name")
    user_prompt = request.form.get("user_prompt", "")

    if not rfp_file:
        return jsonify({"detail": "No RFP file uploaded"}), 400
    # MCP mode: always use default template; ignore UI selection.
    # if not template_name:
    #     return jsonify({"detail": "No template selected"}), 400
    template_name = cfg["default_template"]

    template_path = os.path.join(cfg["datalake_dir"], template_name)
    if not os.path.isfile(template_path):
        return jsonify({"detail": f"Template not found: {template_name}"}), 400

    # Save uploaded RFP to a temp location
    proposal_id = uuid.uuid4().hex[:8]
    ext = os.path.splitext(rfp_file.filename)[1] or ".pdf"
    rfp_path = os.path.join(cfg["output_dir"], f"rfp_{proposal_id}{ext}")
    os.makedirs(cfg["output_dir"], exist_ok=True)
    rfp_file.save(rfp_path)

    try:
        result = run_pipeline(
            client=cfg["client"],
            deployment=cfg["deployment"],
            rfp_path=rfp_path,
            template_path=template_path,
            output_dir=cfg["output_dir"],
            user_prompt=user_prompt,
            embedding_client=cfg.get("embedding_client"),
            embedding_deployment=cfg.get("embedding_deployment"),
        )

        proposals[proposal_id] = result
        return jsonify(
            {
                "proposal_id": proposal_id,
                "message": (
                    f"Generated proposal with {result['sections_mapped']} sections "
                    f"from {result['rfp_sections_found']} RFP sections "
                    f"in {result['timing_seconds']}s"
                ),
                "qa_report": result["qa_report"],
                "proposal": {"slides": result["slides"]},
            }
        )

    except Exception as e:
        traceback.print_exc()
        return jsonify({"detail": f"Pipeline error: {str(e)}"}), 500

    finally:
        # Clean up temp RFP
        if os.path.exists(rfp_path):
            try:
                os.remove(rfp_path)
            except OSError:
                pass


@app.route("/generate_text", methods=["POST"])
def generate_text():
    """MCP entrypoint: accept raw RFP text + user prompt, run pipeline.

    Body JSON: { "rfp_text": "...", "user_prompt": "..." }
    Always uses the default template from cfg["default_template"].
    """
    data = request.get_json(force=True, silent=True) or {}
    rfp_text = data.get("rfp_text", "")
    user_prompt = data.get("user_prompt", "")

    if not rfp_text or not str(rfp_text).strip():
        return jsonify({"detail": "rfp_text is required"}), 400

    template_name = cfg["default_template"]
    template_path = os.path.join(cfg["datalake_dir"], template_name)
    if not os.path.isfile(template_path):
        return jsonify({"detail": f"Default template not found: {template_name}"}), 500

    proposal_id = uuid.uuid4().hex[:8]
    os.makedirs(cfg["output_dir"], exist_ok=True)
    rfp_path = os.path.join(cfg["output_dir"], f"rfp_{proposal_id}.txt")
    with open(rfp_path, "w", encoding="utf-8") as f:
        f.write(rfp_text)

    try:
        result = run_pipeline(
            client=cfg["client"],
            deployment=cfg["deployment"],
            rfp_path=rfp_path,
            template_path=template_path,
            output_dir=cfg["output_dir"],
            user_prompt=user_prompt,
            embedding_client=cfg.get("embedding_client"),
            embedding_deployment=cfg.get("embedding_deployment"),
        )

        proposals[proposal_id] = result
        return jsonify(
            {
                "proposal_id": proposal_id,
                "message": (
                    f"Generated proposal with {result['sections_mapped']} sections "
                    f"from {result['rfp_sections_found']} RFP sections "
                    f"in {result['timing_seconds']}s"
                ),
                "qa_report": result["qa_report"],
                "proposal": {"slides": result["slides"]},
                "download_url": f"/download/{proposal_id}",
                "template_used": template_name,
            }
        )

    except Exception as e:
        traceback.print_exc()
        return jsonify({"detail": f"Pipeline error: {str(e)}"}), 500

    finally:
        if os.path.exists(rfp_path):
            try:
                os.remove(rfp_path)
            except OSError:
                pass


@app.route("/download/<proposal_id>", methods=["GET"])
def download(proposal_id):
    """Download the generated PPTX file."""
    prop = proposals.get(proposal_id)
    if not prop:
        return jsonify({"detail": "Proposal not found"}), 404
    output_path = prop["output_path"]
    if not os.path.isfile(output_path):
        return jsonify({"detail": "File not found on disk"}), 404
    return send_file(
        output_path,
        as_attachment=True,
        download_name=os.path.basename(output_path),
    )


@app.route("/customize", methods=["POST"])
def customize():
    """Re-run pipeline with customization instructions (placeholder for POC)."""
    data = request.get_json(force=True)
    proposal_id = data.get("proposal_id")
    instructions = data.get("instructions", "")

    prop = proposals.get(proposal_id)
    if not prop:
        return jsonify({"detail": "Proposal not found"}), 404

    # For the POC, return the existing proposal with updated message
    return jsonify(
        {
            "proposal_id": proposal_id,
            "message": f"Customization noted: {instructions[:100]}. Re-generation with customizations would go here.",
            "qa_report": prop["qa_report"],
            "proposal": {"slides": prop["slides"]},
        }
    )


# ─── Entry point ─────────────────────────────────────────────────────────────

def _warm_kb_async():
    """Build / sync the persistent KB on a background thread at startup.
    New .pptx uploads are picked up here automatically (manifest tracks
    filename + mtime + size). Only changed files get re-indexed."""
    import threading
    from pipeline import get_kb

    def _run():
        try:
            get_kb(
                cfg["datalake_dir"],
                llm_client=cfg["client"], llm_deployment=cfg["deployment"],
                embedding_client=cfg.get("embedding_client"),
                embedding_deployment=cfg.get("embedding_deployment"),
            )
        except Exception as exc:
            print(f"[KB] warmup crashed: {exc}")

    t = threading.Thread(target=_run, name="kb-warmup", daemon=True)
    t.start()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8501))
    print("=" * 50)
    print(f"  Proposal Assist - starting on http://localhost:{port}")
    print("=" * 50)
    # Pre-warm slide previews for every .pptx in the datalake on a background
    # thread. New uploads picked up automatically (cache keyed by mtime + size).
    warm_layout_cache_async(cfg["datalake_dir"], STATIC_DIR)
    # Build / sync the persistent KB (Chroma + TOC sidecars + embeddings).
    # Incremental — only new / changed templates get re-indexed.
    _warm_kb_async()
    app.run(host="0.0.0.0", port=port, debug=False, threaded=True)
