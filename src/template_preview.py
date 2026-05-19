"""Generate PNG previews of PPTX templates for the gallery UI.

Strategy:
  1. Try PowerPoint COM automation (Windows + PowerPoint installed) - exports the
     first slide of each template to a real PNG image.
  2. Fall back to a Pillow-rendered themed tile that pulls the template's title
     and accent color from the master, so the user still sees a recognizable
     differentiation between templates.

Generated previews are cached in static/template_previews/ keyed by template
filename + mtime.
"""

from __future__ import annotations

import os
import sys
import time
import shutil
import hashlib
import tempfile
import threading
from pathlib import Path
from typing import Optional, Tuple

# PowerPoint COM is single-instance; serialize calls across threads/requests.
_POWERPOINT_LOCK = threading.RLock()

from pptx import Presentation
from pptx.util import Emu


CACHE_DIR_NAME = "template_previews"
LAYOUT_CACHE_DIR_NAME = "template_previews/layouts"
PREVIEW_W = 640
PREVIEW_H = 360
LAYOUT_PREVIEW_W = 480
LAYOUT_PREVIEW_H = 270
MAX_LAYOUTS = 20         # render up to N candidate layouts
MIN_RICH_BYTES = 2500    # below this the PNG is visually empty (filtered out)


# ─── Public API ──────────────────────────────────────────────────────────────

def get_preview_path(template_path: str, static_dir: str) -> Optional[str]:
    """Return absolute path to a cached PNG preview for the given PPTX template.

    Generates the file on demand if missing or stale.
    """
    if not os.path.isfile(template_path):
        return None

    cache_dir = os.path.join(static_dir, CACHE_DIR_NAME)
    os.makedirs(cache_dir, exist_ok=True)

    tpl_stat = os.stat(template_path)
    cache_key = hashlib.md5(
        f"{os.path.basename(template_path)}|{tpl_stat.st_mtime_ns}|{tpl_stat.st_size}".encode()
    ).hexdigest()[:12]
    base = os.path.splitext(os.path.basename(template_path))[0]
    out_png = os.path.join(cache_dir, f"{_slug(base)}__{cache_key}.png")

    if os.path.isfile(out_png) and os.path.getsize(out_png) > 1024:
        return out_png

    # Purge older previews for this template (different mtime)
    for old in Path(cache_dir).glob(f"{_slug(base)}__*.png"):
        if str(old) != out_png:
            try:
                old.unlink()
            except OSError:
                pass

    rendered = _render_with_powerpoint(template_path, out_png) or _render_with_pillow(
        template_path, out_png
    )
    return rendered if rendered and os.path.isfile(rendered) else None


def get_layout_previews(template_path: str, static_dir: str) -> list:
    """Return list of slide-image previews for the template.

    Strategy: copy the template, strip all text from every shape, then export
    each slide as a PNG. The user sees the actual slide designs (images,
    shapes, colors, layout) but with no client content / wording.

    [{ "index": 0, "name": "Slide 1", "png_path": "<abs>", "rel_url": "/static/..." }, ...]
    """
    if not os.path.isfile(template_path):
        return []

    cache_dir = os.path.join(static_dir, LAYOUT_CACHE_DIR_NAME)
    os.makedirs(cache_dir, exist_ok=True)

    tpl_stat = os.stat(template_path)
    cache_key = hashlib.sha256(
        f"{os.path.basename(template_path)}|{tpl_stat.st_mtime_ns}|{tpl_stat.st_size}".encode()
    ).hexdigest()[:12]
    base = _slug(os.path.splitext(os.path.basename(template_path))[0])

    # Count slides and build labels
    try:
        prs = Presentation(template_path)
        total_slides = len(prs.slides)
    except Exception:
        return []

    slide_count = min(total_slides, MAX_LAYOUTS)
    if slide_count <= 0:
        return []
    slide_labels = [f"Slide {i + 1}" for i in range(slide_count)]

    # Purge stale caches for this template
    for old in Path(cache_dir).glob(f"{base}__*__*.png"):
        if cache_key not in old.name:
            try:
                old.unlink()
            except OSError:
                pass

    expected_files = [
        os.path.join(cache_dir, f"{base}__{cache_key}__{idx:02d}.png")
        for idx in range(slide_count)
    ]
    if all(os.path.isfile(f) and os.path.getsize(f) > 1024 for f in expected_files):
        return _format_layout_payload(slide_labels, expected_files, base, cache_key)

    pngs = _render_layouts_with_powerpoint(template_path, slide_labels, expected_files)
    # Pillow fallback intentionally disabled — the user wants real slide
    # exports or nothing at all. Falling through to fake tiles produces a
    # misleading "bluff" preview that hides the real failure.
    return _format_layout_payload(slide_labels, pngs or expected_files, base, cache_key)


def _format_layout_payload(layout_names, png_paths, base, cache_key):
    """Build the JSON payload, keeping ONLY visually-rich previews.

    Tiny PNGs (~1-3 KB) are layouts with no master design — just placeholder
    rectangles. The user only wants the "non empty and visual" ones.
    """
    out = []
    for idx, (name, png) in enumerate(zip(layout_names, png_paths)):
        if not png or not os.path.isfile(png):
            continue
        if os.path.getsize(png) < MIN_RICH_BYTES:
            continue
        rel = f"/static/{LAYOUT_CACHE_DIR_NAME}/{base}__{cache_key}__{idx:02d}.png"
        out.append({"index": idx, "name": name, "png_path": png, "rel_url": rel})
    return out


def warm_layout_cache(datalake_dir: str, static_dir: str, log=print) -> dict:
    """Scan datalake for .pptx files and ensure each has cached slide previews.

    Returns a summary dict {<filename>: <preview-count>}. Cache files live in
    `static/template_previews/layouts/` keyed by (filename, mtime, size) so any
    newly-uploaded or replaced .pptx will be re-rendered automatically the next
    time this runs. Existing cached templates are skipped instantly.
    """
    summary = {}
    if not os.path.isdir(datalake_dir):
        return summary

    pptx_files = sorted(
        p for p in Path(datalake_dir).glob("*.pptx")
        if not p.name.startswith("~$")
    )
    log(f"[preview-cache] scanning {len(pptx_files)} template(s) in {datalake_dir}")

    for pptx in pptx_files:
        _force_kill_powerpoint()  # ensure a clean COM slate before each render
        try:
            t0 = time.time()
            previews = get_layout_previews(str(pptx), static_dir)
            dt = time.time() - t0
            summary[pptx.name] = len(previews)
            log(
                f"[preview-cache] {pptx.name}: {len(previews)} previews "
                f"in {dt:.1f}s"
            )
        except Exception as exc:
            log(f"[preview-cache] {pptx.name}: FAILED -- {exc}")
            summary[pptx.name] = 0
        _force_kill_powerpoint()
        time.sleep(2.5)

    return summary


def _force_kill_powerpoint():
    """Terminate any lingering POWERPNT.EXE processes from prior renders.

    PowerPoint COM sometimes leaves zombie processes that block subsequent
    Presentations.Open() calls. Cheap to call (NOP if no process)."""
    if sys.platform != "win32":
        return
    try:
        import subprocess
        subprocess.run(
            ["taskkill", "/F", "/IM", "POWERPNT.EXE"],
            capture_output=True, timeout=5
        )
    except Exception:
        pass


def warm_layout_cache_async(datalake_dir: str, static_dir: str, log=print):
    """Kick off `warm_layout_cache` on a daemon thread. Server keeps serving;
    previews appear in the gallery as each template finishes rendering."""
    def _run():
        try:
            warm_layout_cache(datalake_dir, static_dir, log=log)
        except Exception as exc:
            log(f"[preview-cache] warmup crashed: {exc}")

    t = threading.Thread(target=_run, name="preview-warmup", daemon=True)
    t.start()
    return t


def get_template_metadata(template_path: str) -> dict:
    """Lightweight template metadata for the gallery card."""
    try:
        prs = Presentation(template_path)
        slides = list(prs.slides)
        sector = _detect_sector_quick(prs, os.path.basename(template_path))
        accent = _detect_accent_color(prs)
        return {
            "slide_count": len(slides),
            "sector": sector,
            "accent_color": accent,
        }
    except Exception:
        return {"slide_count": 0, "sector": "General", "accent_color": "#FFE600"}


# ─── PowerPoint COM path ─────────────────────────────────────────────────────

def _create_powerpoint_with_retry(comtypes_client, attempts: int = 12):
    """COM 'Call was rejected by callee' (-2147418111) happens when PowerPoint is
    still starting up or closing from a prior request. Retry with backoff."""
    last_exc = None
    for i in range(attempts):
        try:
            pp = comtypes_client.CreateObject("PowerPoint.Application")
            return pp
        except Exception as exc:
            last_exc = exc
            time.sleep(0.5 + 0.4 * i)
    if last_exc:
        raise last_exc
    return None


def _open_presentation_with_retry(powerpoint, abs_path: str, read_only: bool, attempts: int = 10):
    last_exc = None
    for i in range(attempts):
        try:
            return powerpoint.Presentations.Open(
                abs_path,
                ReadOnly=-1 if read_only else 0,
                Untitled=0,
                WithWindow=0,
            )
        except Exception as exc:
            last_exc = exc
            time.sleep(0.5 + 0.4 * i)
    if last_exc:
        raise last_exc
    return None


def _render_with_powerpoint(template_path: str, out_png: str) -> Optional[str]:
    if sys.platform != "win32":
        return None
    try:
        import comtypes
        import comtypes.client  # type: ignore
    except Exception:
        return None

    workdir = tempfile.mkdtemp(prefix="ppx_preview_")
    work_path = os.path.join(workdir, os.path.basename(template_path))
    try:
        shutil.copy2(template_path, work_path)
    except Exception:
        shutil.rmtree(workdir, ignore_errors=True)
        return None

    powerpoint = None
    presentation = None
    try:
        comtypes.CoInitialize()  # per-thread COM init for Flask workers
    except OSError:
        pass
    with _POWERPOINT_LOCK:
        try:
            comtypes.client.gen_dir = None
            powerpoint = _create_powerpoint_with_retry(comtypes.client)
            try:
                powerpoint.DisplayAlerts = 0
            except Exception:
                pass
            try:
                powerpoint.Visible = 1
            except Exception:
                pass

            presentation = _open_presentation_with_retry(
                powerpoint, os.path.abspath(work_path), read_only=True
            )

            if presentation.Slides.Count < 1:
                return None
            slide = presentation.Slides(1)
            slide.Export(os.path.abspath(out_png), "PNG", PREVIEW_W, PREVIEW_H)
            if os.path.isfile(out_png) and os.path.getsize(out_png) > 1024:
                return out_png
            return None
        except Exception as exc:
            print(f"  [preview] PowerPoint COM failed for {os.path.basename(template_path)}: {exc}")
            return None
        finally:
            try:
                if presentation is not None:
                    presentation.Close()
            except Exception:
                pass
            try:
                if powerpoint is not None:
                    powerpoint.Quit()
            except Exception:
                pass
            shutil.rmtree(workdir, ignore_errors=True)


# ─── PowerPoint COM: per-layout rendering ────────────────────────────────────

def _build_textless_pptx(template_path: str, workdir: str):
    """Copy the template and clear every run of text on every slide, keeping
    images / shapes / colors / layout intact. Returns the path on success."""
    try:
        from pptx import Presentation
    except Exception:
        return None

    out_path = os.path.join(workdir, "_textless.pptx")
    try:
        shutil.copy2(template_path, out_path)
        prs = Presentation(out_path)
        for slide in prs.slides:
            _strip_text_from_shapes(slide.shapes)
        prs.save(out_path)
        return out_path
    except Exception as exc:
        print(f"  [layout-preview] text-stripping failed: {exc}")
        return None


def _strip_text_from_shapes(shapes):
    """Recursively wipe text content from all shapes (groups, tables, frames)."""
    for shape in shapes:
        try:
            if getattr(shape, "shape_type", None) == 6:  # MSO_SHAPE_TYPE.GROUP
                _strip_text_from_shapes(shape.shapes)
                continue
        except Exception:
            pass
        try:
            if getattr(shape, "has_text_frame", False):
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        try:
                            run.text = ""
                        except Exception:
                            continue
        except Exception:
            pass
        try:
            if getattr(shape, "has_table", False):
                for row in shape.table.rows:
                    for cell in row.cells:
                        for para in cell.text_frame.paragraphs:
                            for run in para.runs:
                                try:
                                    run.text = ""
                                except Exception:
                                    continue
        except Exception:
            pass


def _render_layouts_with_powerpoint(template_path, layout_names, out_files):
    """Render layout previews via Slide Master-style empty slides.

    1. Build a stripped .pptx (python-pptx, no client content).
    2. Open in PowerPoint and export each slide.
    """
    if sys.platform != "win32":
        return None
    try:
        import comtypes
        import comtypes.client  # type: ignore
    except Exception:
        return None

    workdir = tempfile.mkdtemp(prefix="ppx_layouts_")
    textless_path = _build_textless_pptx(template_path, workdir)
    if not textless_path:
        shutil.rmtree(workdir, ignore_errors=True)
        return None

    rendered = [None] * len(layout_names)
    powerpoint = None
    presentation = None
    try:
        comtypes.CoInitialize()  # per-thread COM init for Flask workers
    except OSError:
        pass
    with _POWERPOINT_LOCK:
        try:
            comtypes.client.gen_dir = None
            powerpoint = _create_powerpoint_with_retry(comtypes.client)
            try:
                powerpoint.DisplayAlerts = 0
            except Exception:
                pass
            try:
                powerpoint.Visible = 1
            except Exception:
                pass

            presentation = _open_presentation_with_retry(
                powerpoint, os.path.abspath(textless_path), read_only=True
            )

            slide_count = presentation.Slides.Count
            for idx in range(len(layout_names)):
                slide_num = idx + 1
                if slide_num > slide_count:
                    break
                out = os.path.abspath(out_files[idx])
                os.makedirs(os.path.dirname(out), exist_ok=True)
                try:
                    presentation.Slides(slide_num).Export(
                        out, "PNG", LAYOUT_PREVIEW_W, LAYOUT_PREVIEW_H
                    )
                    if os.path.isfile(out) and os.path.getsize(out) > 512:
                        rendered[idx] = out
                except Exception as exc:
                    print(
                        f"  [layout-preview] export slide {slide_num} failed: {exc}"
                    )
            return rendered
        except Exception as exc:
            print(f"  [layout-preview] PowerPoint COM failed: {exc}")
            return None
        finally:
            try:
                if presentation is not None:
                    presentation.Close()
            except Exception:
                pass
            try:
                if powerpoint is not None:
                    powerpoint.Quit()
            except Exception:
                pass
            shutil.rmtree(workdir, ignore_errors=True)


def _render_layouts_with_pillow(template_path, layout_names, out_files):
    try:
        from PIL import Image, ImageDraw
    except Exception:
        return None

    try:
        prs = Presentation(template_path)
        accent = _detect_accent_color(prs)
        bg = _detect_background_color(prs) or "#1A1A2E"
    except Exception:
        accent = "#FFE600"
        bg = "#1A1A2E"

    bg_rgb = _hex_to_rgb(bg)
    accent_rgb = _hex_to_rgb(accent)
    rendered = []
    for idx, name in enumerate(layout_names):
        out = out_files[idx]
        os.makedirs(os.path.dirname(out), exist_ok=True)
        img = Image.new("RGB", (LAYOUT_PREVIEW_W, LAYOUT_PREVIEW_H), bg_rgb)
        draw = ImageDraw.Draw(img)
        draw.rectangle((0, 0, 8, LAYOUT_PREVIEW_H), fill=accent_rgb)
        title_font = _load_font(20, bold=True)
        sub_font = _load_font(12)
        draw.text((28, 28), name[:38], fill=(255, 255, 255), font=title_font)
        draw.text((28, 56), "Master layout", fill=(220, 220, 230), font=sub_font)
        # Stripe pattern to vary previews
        for i in range(3):
            x = 28 + i * 140
            draw.rounded_rectangle(
                (x, 130, x + 110, 220), radius=6,
                fill=_shade(bg_rgb, 1.35),
                outline=accent_rgb, width=2,
            )
        img.save(out, "PNG")
        rendered.append(out if os.path.isfile(out) else None)
    return rendered


# ─── Pillow fallback ─────────────────────────────────────────────────────────

def _render_with_pillow(template_path: str, out_png: str) -> Optional[str]:
    try:
        from PIL import Image, ImageDraw, ImageFont
    except Exception:
        return None

    try:
        prs = Presentation(template_path)
        accent = _detect_accent_color(prs)
        bg = _detect_background_color(prs) or "#1A1A2E"
        title = _detect_first_title(prs) or os.path.splitext(
            os.path.basename(template_path)
        )[0]
    except Exception:
        accent = "#FFE600"
        bg = "#1A1A2E"
        title = os.path.splitext(os.path.basename(template_path))[0]

    img = Image.new("RGB", (PREVIEW_W, PREVIEW_H), _hex_to_rgb(bg))
    draw = ImageDraw.Draw(img)

    # Accent ribbon on the left
    draw.rectangle((0, 0, 14, PREVIEW_H), fill=_hex_to_rgb(accent))

    # Soft diagonal gradient by overlay
    overlay = Image.new("RGB", (PREVIEW_W, PREVIEW_H), _hex_to_rgb(bg))
    overlay_draw = ImageDraw.Draw(overlay)
    overlay_draw.polygon(
        [(0, PREVIEW_H), (PREVIEW_W, PREVIEW_H), (PREVIEW_W, 80)],
        fill=_shade(_hex_to_rgb(bg), 0.85),
    )
    img = Image.blend(img, overlay, 0.45)
    draw = ImageDraw.Draw(img)

    # Faux title bar
    title_font = _load_font(34)
    sub_font = _load_font(16)
    short = title[:48] + ("..." if len(title) > 48 else "")
    draw.text((48, 110), short, fill=(255, 255, 255), font=title_font)
    draw.text(
        (48, 162),
        "Template Theme Preview",
        fill=(220, 220, 230),
        font=sub_font,
    )

    # Three faux content blocks
    block_y = 220
    for i in range(3):
        x0 = 48 + i * 180
        draw.rounded_rectangle(
            (x0, block_y, x0 + 150, block_y + 80),
            radius=6,
            fill=_shade(_hex_to_rgb(bg), 1.4),
            outline=_hex_to_rgb(accent),
            width=2,
        )

    # Accent chip
    draw.rectangle((PREVIEW_W - 110, 22, PREVIEW_W - 22, 50), fill=_hex_to_rgb(accent))
    draw.text(
        (PREVIEW_W - 100, 26),
        "Theme",
        fill=(0, 0, 0),
        font=_load_font(14, bold=True),
    )

    img.save(out_png, "PNG")
    return out_png if os.path.isfile(out_png) else None


# ─── Small helpers ───────────────────────────────────────────────────────────

def _slug(name: str) -> str:
    return "".join(c if c.isalnum() else "_" for c in name).strip("_") or "tpl"


def _hex_to_rgb(hex_color: str) -> Tuple[int, int, int]:
    h = hex_color.lstrip("#")
    if len(h) == 3:
        h = "".join(c * 2 for c in h)
    try:
        return tuple(int(h[i : i + 2], 16) for i in (0, 2, 4))  # type: ignore
    except Exception:
        return (26, 26, 46)


def _shade(rgb: Tuple[int, int, int], factor: float) -> Tuple[int, int, int]:
    return tuple(max(0, min(255, int(c * factor))) for c in rgb)  # type: ignore


def _load_font(size: int, bold: bool = False):
    from PIL import ImageFont

    candidates = [
        "C:\\Windows\\Fonts\\segoeuib.ttf" if bold else "C:\\Windows\\Fonts\\segoeui.ttf",
        "C:\\Windows\\Fonts\\arialbd.ttf" if bold else "C:\\Windows\\Fonts\\arial.ttf",
    ]
    for path in candidates:
        try:
            return ImageFont.truetype(path, size=size)
        except Exception:
            continue
    return ImageFont.load_default()


def _detect_first_title(prs: Presentation) -> Optional[str]:
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame and "title" in shape.name.lower():
                txt = (shape.text_frame.text or "").strip()
                if txt:
                    return txt[:80]
        for shape in slide.shapes:
            if shape.has_text_frame:
                txt = (shape.text_frame.text or "").strip()
                if txt:
                    return txt[:80]
        break
    return None


def _detect_accent_color(prs: Presentation) -> str:
    """Extract a reasonable accent color from theme XML."""
    try:
        masters = list(prs.slide_masters)
        if not masters:
            return "#FFE600"
        theme = masters[0].element.getroottree().getroot()
        ns = {
            "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        }
        # Look at first <a:srgbClr> under master fills
        for el in masters[0].element.iter():
            tag = el.tag.split("}", 1)[-1] if "}" in el.tag else el.tag
            if tag == "srgbClr" and el.get("val"):
                val = el.get("val", "")
                if len(val) == 6 and not val.lower().startswith(("ffffff", "000000")):
                    return f"#{val.upper()}"
    except Exception:
        pass
    return "#FFE600"


def _detect_background_color(prs: Presentation) -> Optional[str]:
    try:
        masters = list(prs.slide_masters)
        if not masters:
            return None
        for el in masters[0].element.iter():
            tag = el.tag.split("}", 1)[-1] if "}" in el.tag else el.tag
            if tag == "srgbClr" and el.get("val"):
                v = el.get("val", "").lower()
                if v not in ("ffffff", "000000"):
                    return f"#{v.upper()}"
    except Exception:
        return None
    return None


def _detect_sector_quick(prs: Presentation, filename: str) -> str:
    fl = filename.lower()
    if "health" in fl:
        return "Healthcare"
    if "oil" in fl or "gas" in fl:
        return "Oil & Gas"
    if "tech" in fl or "digital" in fl:
        return "Technology"
    if "infra" in fl:
        return "Infrastructure"
    return "General"
