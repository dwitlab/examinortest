"""DOCX -> LaTeX converter web app (Flask + Pandoc + SQLite).

Pandoc does the heavy lifting (Word equations/OMML -> LaTeX, tables, images,
lists, footnotes). This app adds upload validation, LaTeX post-processing,
conversion stats/warnings, and a persistent per-browser history stored in
SQLite (LaTeX text + extracted images as BLOBs).

Requires the `pandoc` binary on PATH (2.11+ works; newer is better).
"""
from __future__ import annotations

import io
import mimetypes
import os
import re
import secrets
import shutil
import subprocess
import tempfile
import uuid
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from flask import (Flask, Response, abort, jsonify, render_template, request,
                   send_file, session)
from common import pandoc_available, pandoc_supports, validate_docx
from extensions import db

# ----------------------------------------------------------------- config
MAX_UPLOAD_MB = int(os.environ.get("MAX_UPLOAD_MB", 25))
MAX_UNZIPPED_MB = 300          # zip-bomb guard for the .docx container
MAX_HISTORY = int(os.environ.get("MAX_HISTORY", 50))   # per browser; oldest pruned
PANDOC_TIMEOUT = 90

DOC_CLASSES = {"article", "report", "book"}
FONT_SIZES = {"10pt", "11pt", "12pt"}
WRAP_MODES = {"none", "auto", "preserve"}
TRACK_MODES = {"accept", "reject", "all"}
LATEX_IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".pdf"}   # what pdfLaTeX can embed

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024
Path(app.instance_path).mkdir(parents=True, exist_ok=True)
app.config["SQLALCHEMY_DATABASE_URI"] = os.environ.get(
    "DATABASE_URL", "sqlite:///" + str(Path(app.instance_path) / "docx2latex.db"))
# Exam text/answers/results live in the main DB; exam images live in their own file.
app.config["SQLALCHEMY_BINDS"] = {"media": os.environ.get(
    "MEDIA_DATABASE_URL", "sqlite:///" + str(Path(app.instance_path) / "media.db"))}


def _secret_key() -> str:
    """Stable secret for the session cookie: env var, else a file in instance/."""
    if os.environ.get("SECRET_KEY"):
        return os.environ["SECRET_KEY"]
    f = Path(app.instance_path) / "secret_key"
    if not f.exists():
        f.write_text(secrets.token_hex(32))
    return f.read_text().strip()


app.secret_key = _secret_key()
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["PERMANENT_SESSION_LIFETIME"] = 60 * 60 * 24 * 365

db.init_app(app)
from exam import bp as exam_bp  # noqa: E402  (imported here so exam tables exist before create_all)
app.register_blueprint(exam_bp)


# ----------------------------------------------------------------- models
def utcnow():
    return datetime.now(timezone.utc)


class Conversion(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    owner = db.Column(db.String(32), index=True, nullable=False)   # anonymous browser id
    filename = db.Column(db.String(255), nullable=False)
    latex = db.Column(db.Text, nullable=False)
    stats = db.Column(db.JSON, nullable=False)
    warnings = db.Column(db.JSON, nullable=False)
    created = db.Column(db.DateTime, default=utcnow, nullable=False)
    images = db.relationship("Image", backref="conversion",
                             cascade="all, delete-orphan", lazy="select")

    @property
    def stem(self) -> str:
        return re.sub(r"[^\w.-]+", "_", Path(self.filename).stem)[:80] or "document"


class Image(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    conversion_id = db.Column(db.ForeignKey("conversion.id"), index=True, nullable=False)
    path = db.Column(db.String(300), nullable=False)    # as referenced in the .tex, e.g. media/image1.jpg
    mimetype = db.Column(db.String(60))
    data = db.Column(db.LargeBinary, nullable=False)


with app.app_context():
    db.create_all()


# ---------------------------------------------------------------- helpers
def current_owner() -> str:
    if "owner" not in session:
        session["owner"] = uuid.uuid4().hex
    session.permanent = True
    return session["owner"]


def get_owned(conv_id: int) -> Conversion:
    conv = db.session.get(Conversion, conv_id)
    if conv is None or conv.owner != current_owner():
        abort(404)          # same answer for "missing" and "not yours"
    return conv


def run_pandoc(docx: Path, job_dir: Path, opts: dict) -> tuple[str, str]:
    cmd = [
        "pandoc", docx.name, "-f", "docx", "-t", "latex",
        "-o", "main.tex",
        "--extract-media=.",                 # images land in ./media/
        f"--wrap={opts['wrap']}",
        f"--track-changes={opts['track']}",
    ]
    if pandoc_supports("--sandbox"):         # pandoc >= 2.15: reader/writer can't touch other files
        cmd.append("--sandbox")
    if opts["standalone"]:
        cmd += ["-s", "-V", f"documentclass={opts['docclass']}",
                "-V", f"fontsize={opts['fontsize']}",
                "-V", f"geometry:margin={opts['margin']}"]
    try:
        proc = subprocess.run(cmd, cwd=job_dir, capture_output=True, text=True,
                              timeout=PANDOC_TIMEOUT)
    except subprocess.TimeoutExpired:
        raise RuntimeError("Conversion timed out - the document may be too complex.")
    if proc.returncode != 0:
        raise RuntimeError("Pandoc failed: " + (proc.stderr.strip() or "unknown error"))
    return (job_dir / "main.tex").read_text(encoding="utf-8"), proc.stderr.strip()


_IMG_RE = re.compile(r"\\includegraphics\[([^\]]*)\]\{([^}]*)\}")


_INLINE_MATH_RE = re.compile(r"\\\((.*?)\\\)", re.S)
_TABLE_RE = re.compile(r"\\begin\{longtable\}.*?\\end\{longtable\}", re.S)
_ANY_IMG_RE = re.compile(r"\\includegraphics\[[^\]]*\]\{[^}]*\}")
_JUNK_BREAK_RE = re.compile(r"\\textbf\{\\hfill\\break\s*\}[ \t]*\n?")
_CAPS_HEADING_RE = re.compile(r"^\\textbf\{([A-Z][A-Z0-9 ,.:&/()\-]{5,})\}[ \t]*$", re.M)


def postprocess(tex: str, opts: dict) -> str:
    """Make pandoc's output nicer for documents authored in Word."""
    # 1. Images: keep Word's sizes but never overflow the text width.
    def fix_img(m: re.Match) -> str:
        o = m.group(1)
        if "width=" in o and "height=" in o:
            o += ",keepaspectratio,max width=\\linewidth"
        return f"\\includegraphics[{o}]{{{m.group(2)}}}"
    tex = _IMG_RE.sub(fix_img, tex)

    # 2. Images inside tables: top-align them so neighbouring cells (labels like
    #    "(A)") sit beside the top of the picture instead of its bottom edge.
    tex = _TABLE_RE.sub(
        lambda t: _ANY_IMG_RE.sub(lambda i: "\\raisebox{-\\height}{" + i.group(0) + "}", t.group(0)),
        tex)

    # 3. Inline fractions render at tiny script size; use \dfrac (amsmath).
    if opts.get("bigfrac", True):
        tex = _INLINE_MATH_RE.sub(
            lambda m: "\\(" + m.group(1).replace("\\frac{", "\\dfrac{") + "\\)", tex)

    # 4. Word's empty bold line-breaks produce junk `\textbf{\hfill\break}` lines.
    tex = _JUNK_BREAK_RE.sub("", tex)

    # 5. Documents without Word heading styles: promote ALL-CAPS bold lines.
    if opts.get("headings", True):
        tex = _CAPS_HEADING_RE.sub(lambda m: "\\section*{" + m.group(1) + "}", tex)

    tex = re.sub(r"\n{3,}", "\n\n", tex)

    # 6. Standalone: adjustbox gives us `max width`; amssymb covers extra symbols.
    if opts["standalone"] and "\\begin{document}" in tex:
        extra = ("\\usepackage[export]{adjustbox}\n"
                 "\\usepackage{amssymb}\n")
        tex = tex.replace("\\begin{document}", extra + "\\begin{document}", 1)
    return tex


def collect_report(tex: str, job_dir: Path, pandoc_stderr: str) -> tuple[dict, list[str]]:
    stats = {
        "equations": len(re.findall(r"\\\(|\\\[", tex)),
        "images": len(re.findall(r"\\includegraphics", tex)),
        "tables": len(re.findall(r"\\begin\{longtable\}", tex)),
        "characters": len(tex),
    }
    warnings: list[str] = []
    media = job_dir / "media"
    bad = sorted(p.name for p in media.rglob("*")
                 if p.is_file() and p.suffix.lower() not in LATEX_IMAGE_EXTS) if media.exists() else []
    if bad:
        warnings.append(
            f"{len(bad)} image(s) are in a format pdfLaTeX cannot embed "
            f"({', '.join(bad[:5])}{'...' if len(bad) > 5 else ''}). "
            "Convert them to PNG/PDF (e.g. `inkscape` or LibreOffice for EMF/WMF).")
    if pandoc_stderr:
        warnings += [ln for ln in pandoc_stderr.splitlines() if ln.strip()][:10]
    return stats, warnings


def parse_options(form) -> dict:
    def pick(name, allowed, default):
        v = form.get(name, default)
        return v if v in allowed else default
    margin = form.get("margin", "1in")
    if not re.fullmatch(r"\d{1,2}(\.\d+)?(in|cm|mm|pt)", margin):
        margin = "1in"
    return {
        "standalone": form.get("standalone", "true") == "true",
        "docclass": pick("docclass", DOC_CLASSES, "article"),
        "fontsize": pick("fontsize", FONT_SIZES, "11pt"),
        "wrap": pick("wrap", WRAP_MODES, "auto"),
        "track": pick("track", TRACK_MODES, "accept"),
        "margin": margin,
        "bigfrac": form.get("bigfrac", "true") == "true",
        "headings": form.get("headings", "true") == "true",
    }


def conversion_json(c: Conversion, with_latex: bool = True) -> dict:
    d = {"id": c.id, "name": c.stem, "filename": c.filename,
         "created": c.created.replace(tzinfo=timezone.utc).isoformat(),
         "stats": c.stats, "warnings": c.warnings}
    if with_latex:
        d["latex"] = c.latex
    return d


def prune_history(owner: str) -> None:
    old = (Conversion.query.filter_by(owner=owner)
           .order_by(Conversion.created.desc(), Conversion.id.desc())
           .offset(MAX_HISTORY).all())
    for c in old:
        db.session.delete(c)
    if old:
        db.session.commit()


# ----------------------------------------------------------------- routes
@app.get("/converter")
def converter():
    current_owner()
    return render_template("index.html", max_mb=MAX_UPLOAD_MB,
                           pandoc_ok=pandoc_available())


@app.get("/health")
def health():
    return jsonify(status="ok", pandoc=pandoc_available())


@app.post("/api/convert")
def convert():
    if not pandoc_available():
        return jsonify(error="Pandoc is not installed on the server."), 503
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify(error="No file uploaded."), 400
    if not f.filename.lower().endswith(".docx"):
        return jsonify(error="Please upload a .docx file."), 400

    owner = current_owner()
    try:
        with tempfile.TemporaryDirectory(prefix="docx2latex_") as tmp:
            job_dir = Path(tmp)
            docx = job_dir / "source.docx"      # never trust the client filename on disk
            f.save(docx)
            validate_docx(docx)
            opts = parse_options(request.form)
            tex, stderr = run_pandoc(docx, job_dir, opts)
            tex = postprocess(tex, opts)
            stats, warnings = collect_report(tex, job_dir, stderr)

            conv = Conversion(owner=owner, filename=Path(f.filename).name[:255],
                              latex=tex, stats=stats, warnings=warnings)
            media = job_dir / "media"
            if media.exists():
                for p in sorted(media.rglob("*")):
                    if p.is_file():
                        rel = p.relative_to(job_dir).as_posix()
                        conv.images.append(Image(
                            path=rel, data=p.read_bytes(),
                            mimetype=mimetypes.guess_type(p.name)[0] or "application/octet-stream"))
            db.session.add(conv)
            db.session.commit()
    except ValueError as e:
        return jsonify(error=str(e)), 400
    except RuntimeError as e:
        return jsonify(error=str(e)), 500

    prune_history(owner)
    return jsonify(conversion_json(conv))


@app.get("/api/history")
def history():
    rows = (Conversion.query.filter_by(owner=current_owner())
            .order_by(Conversion.created.desc(), Conversion.id.desc()).all())
    return jsonify([conversion_json(c, with_latex=False) for c in rows])


@app.get("/api/conversions/<int:conv_id>")
def get_conversion(conv_id: int):
    return jsonify(conversion_json(get_owned(conv_id)))


@app.delete("/api/conversions/<int:conv_id>")
def delete_conversion(conv_id: int):
    db.session.delete(get_owned(conv_id))
    db.session.commit()
    return jsonify(ok=True)


@app.get("/download/<int:conv_id>/<kind>")
def download(conv_id: int, kind: str):
    if kind not in {"tex", "zip"}:
        abort(404)
    conv = get_owned(conv_id)
    if kind == "tex":
        return Response(conv.latex, mimetype="application/x-tex", headers={
            "Content-Disposition": f'attachment; filename="{conv.stem}.tex"'})
    buf = io.BytesIO()                      # ZIP is assembled on demand from the DB
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("main.tex", conv.latex)
        for img in conv.images:
            z.writestr(img.path, img.data)
    buf.seek(0)
    return send_file(buf, mimetype="application/zip", as_attachment=True,
                     download_name=f"{conv.stem}_latex.zip")


@app.get("/image/<int:image_id>")
def get_image(image_id: int):
    img = db.session.get(Image, image_id)
    if img is None or img.conversion.owner != current_owner():
        abort(404)
    return Response(img.data, mimetype=img.mimetype)


@app.errorhandler(413)
def too_large(_):
    return jsonify(error=f"File exceeds the {MAX_UPLOAD_MB} MB limit."), 413


if __name__ == "__main__":
    app.run(debug=os.environ.get("FLASK_DEBUG") == "1", port=5000)
