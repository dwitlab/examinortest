"""Online exam module.

Examiner uploads a .docx question paper -> it is parsed into questions (type,
answer key, images, equations) -> examiner reviews/edits -> publishes -> students
join with a code, take a timed exam, and are graded automatically.
"""
from __future__ import annotations

import csv
import io
import mimetypes
import os
import re
import secrets
import subprocess
import tempfile
import uuid
from datetime import datetime, timedelta, timezone
from functools import wraps
from pathlib import Path

import lxml.html as LH
from flask import (Blueprint, Response, abort, current_app, flash, jsonify,
                   redirect, render_template, request, session, url_for)

from common import pandoc_available, pandoc_supports, validate_docx
from extensions import db

bp = Blueprint("exam", __name__)

GRACE_SECONDS = 5                       # network slack when saving near the deadline
LETTERS = "ABCD"
DEFAULT_MARKS = {"single": (4.0, 1.0), "multiple": (4.0, 2.0), "integer": (4.0, 0.0)}
TYPE_LABELS = {"single": "Single correct", "multiple": "Multiple correct", "integer": "Numerical"}


def now() -> datetime:
    """Naive UTC (SQLite stores naive datetimes)."""
    return datetime.now(timezone.utc).replace(tzinfo=None)


def iso(dt: datetime | None) -> str | None:
    return dt.isoformat() + "Z" if dt else None


# ----------------------------------------------------------------- models
class Exam(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(200), nullable=False)
    code = db.Column(db.String(8), unique=True, index=True, nullable=False)
    duration = db.Column(db.Integer, nullable=False, default=60)       # minutes
    published = db.Column(db.Boolean, default=False, nullable=False)
    show_answers = db.Column(db.Boolean, default=True, nullable=False)
    created = db.Column(db.DateTime, default=now, nullable=False)
    questions = db.relationship("ExamQuestion", order_by="ExamQuestion.position",
                                cascade="all, delete-orphan", backref="exam")
    attempts = db.relationship("Attempt", cascade="all, delete-orphan", backref="exam")

    @property
    def total_marks(self) -> float:
        return sum(q.marks for q in self.questions)


class ExamQuestion(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    exam_id = db.Column(db.ForeignKey("exam.id"), index=True, nullable=False)
    position = db.Column(db.Integer, nullable=False)
    orig_no = db.Column(db.Integer)                              # number printed in the source paper
    section = db.Column(db.String(300), default="")
    qtype = db.Column(db.String(10), nullable=False)             # single | multiple | integer
    label_style = db.Column(db.String(5), default="alpha")       # how options are labelled: alpha (A) | num (1)
    context = db.Column(db.Text, default="")                     # shared passage (comprehension)
    html = db.Column(db.Text, nullable=False)
    explanation = db.Column(db.Text, default="")
    answer = db.Column(db.String(60), default="")                # "C", "A,C,D" or "9.30"
    marks = db.Column(db.Float, default=4.0)
    negative = db.Column(db.Float, default=0.0)

    def label(self, letter: str) -> str:
        return str(LETTERS.index(letter) + 1) if self.label_style == "num" else letter

    def answer_display(self) -> str:
        if self.qtype == "integer" or not self.answer:
            return self.answer or "-"
        return ", ".join(self.label(a) for a in self.answer.split(","))


class ExamAsset(db.Model):
    """Exam images. Stored in a SEPARATE SQLite file (media.db) from the question text (exam.db).
    SQLite cannot enforce foreign keys across files, so exam_id is a plain indexed integer and
    deleting an exam removes its assets explicitly (see admin_delete)."""
    __bind_key__ = "media"
    id = db.Column(db.Integer, primary_key=True)
    exam_id = db.Column(db.Integer, index=True, nullable=False)
    name = db.Column(db.String(200), nullable=False)
    mimetype = db.Column(db.String(60))
    data = db.Column(db.LargeBinary, nullable=False)


class Attempt(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    exam_id = db.Column(db.ForeignKey("exam.id"), index=True, nullable=False)
    token = db.Column(db.String(32), unique=True, index=True, nullable=False)
    name = db.Column(db.String(120), nullable=False)
    roll = db.Column(db.String(60), nullable=False)
    started = db.Column(db.DateTime, default=now, nullable=False)
    submitted = db.Column(db.DateTime)
    answers = db.Column(db.JSON, default=dict, nullable=False)
    score = db.Column(db.Float)
    total = db.Column(db.Float)
    details = db.Column(db.JSON)

    @property
    def deadline(self) -> datetime:
        return self.started + timedelta(minutes=self.exam.duration)

    @property
    def remaining(self) -> int:
        return max(0, int((self.deadline - now()).total_seconds()))


# ------------------------------------------------------------ DOCX parsing
Q_START = re.compile(r"^\s*(\d{1,3})\s*[.)]\s+\S")
ANS = re.compile(r"^\s*(?:ans(?:wer)?|key)\s*(?:[.:\-–]|(?=\())\s*(.*)$", re.I | re.S)
SOL = re.compile(r"^\s*(?:sol(?:ution)?|explanation)\b", re.I)
LEAD_NUM = re.compile(r"^<p>\s*(?:<strong>)?\s*\d{1,3}\s*[.)]\s*(?:</strong>)?\s*")
CHOICE = re.compile(r"^\(?\s*([A-Da-d1-4](?:\s*[,&/;]\s*[A-Da-d1-4])*)\s*\)?(?=$|[\s\\])")
NUMBER = re.compile(r"^\(?\s*(-?\d+(?:\.\d+)?)\s*\)?")


def docx_to_html(docx: Path, job_dir: Path) -> str:
    cmd = ["pandoc", docx.name, "-f", "docx", "-t", "html", "--mathjax", "--wrap=none",
           "--extract-media=.", "--track-changes=accept"]
    if pandoc_supports("--sandbox"):
        cmd.append("--sandbox")
    try:
        p = subprocess.run(cmd, cwd=job_dir, capture_output=True, text=True, timeout=90)
    except subprocess.TimeoutExpired:
        raise RuntimeError("Conversion timed out.")
    if p.returncode != 0:
        raise RuntimeError("Pandoc failed: " + (p.stderr.strip() or "unknown error"))
    return p.stdout


def type_from_heading(h: str) -> str | None:
    u = h.upper()
    if re.search(r"INTEGER|NUMERIC", u):
        return "integer"
    if "MULTIPLE" in u:
        return "multiple"
    if re.search(r"SINGLE|SCQ|MATRIX|MATCH", u):
        return "single"
    return None


def normalize_key(qtype: str, raw: str) -> str:
    """Canonical answer key: 'C', 'A,C,D' (letters) or '9.30' (numeric)."""
    raw = (raw or "").strip()
    if not raw:
        return ""
    if qtype == "integer":
        m = NUMBER.match(raw)
        return m.group(1) if m else ""
    m = CHOICE.match(raw)
    if not m:
        return ""
    letters = {(LETTERS["1234".index(t)] if t in "1234" else t.upper())
               for t in re.findall(r"[A-Da-d1-4]", m.group(1))}
    return ",".join(sorted(letters))


def parse_answer(raw: str, hint: str | None) -> tuple[str, str]:
    """-> (qtype, key). Heading hint wins; otherwise infer from the key itself."""
    raw = raw.strip()
    if hint == "integer":
        return "integer", normalize_key("integer", raw)
    key = normalize_key("single", raw)
    if key:
        return ("multiple" if "," in key else (hint if hint in ("single", "multiple") else "single")), key
    key = normalize_key("integer", raw)
    return ("integer", key) if key else (hint or "single", "")


def html_of(els) -> str:
    return "".join(LH.tostring(e, encoding="unicode", with_tail=False) for e in els)


def parse_questions(html: str) -> tuple[list[dict], list[str]]:
    root = LH.fragment_fromstring(html, create_parent="div")
    for s in root.iter("span"):                       # readable inline fractions
        if "math" in (s.get("class") or "") and "inline" in (s.get("class") or "") and s.text:
            s.text = s.text.replace("\\frac{", "\\dfrac{")

    out: list[dict] = []
    section, hint = "", None
    ctx: list = []            # shared passage elements
    ctx_open = False          # still collecting a passage
    cur: dict | None = None
    in_answer = False

    def finish():
        nonlocal cur
        if cur is not None:
            out.append(cur)
        cur = None

    for el in root:
        text = " ".join(el.text_content().split())
        has_img = el.find(".//img") is not None
        if not text and not has_img:
            continue

        # Section heading: an all-caps line made of real words (not "(D) (A) -> (P, S)" option text)
        is_head = (el.tag == "p" and not has_img and el.find(".//span") is None
                   and text[0].isalpha() and text == text.upper()
                   and len(re.findall(r"[A-Z]{4,}", text)) >= 2 and not ANS.match(text))
        if is_head:
            finish()
            section, ctx, ctx_open, in_answer = text, [], False, False
            hint = type_from_heading(text) or hint
            continue

        if el.tag == "p" and Q_START.match(text):
            finish()
            cur = {"orig": int(Q_START.match(text).group(1)), "section": section,
                   "els": [el], "expl": [], "raw": None,
                   "hint": hint, "ctx": list(ctx)}
            ctx_open, in_answer = False, False
            continue

        m = ANS.match(text) if el.tag == "p" else None
        if m and cur is not None and not in_answer:
            raw = m.group(1).strip()
            cur["raw"] = raw
            in_answer = True
            km = CHOICE.match(raw) or NUMBER.match(raw)
            if km and raw[km.end():].strip():
                cur["expl"].append(el)               # mapping/working written on the Ans line
            continue

        if cur is not None and in_answer:
            if SOL.match(text):
                cur["expl"].append(el)
                continue
            if el.tag == "p" and Q_START.match(text) is None and not ctx_open:
                finish()                                # passage for the next questions
                ctx, ctx_open = [el], True
                continue
        if cur is None:
            if not ctx_open:
                ctx, ctx_open = [], True
            ctx.append(el)
            continue
        (cur["expl"] if in_answer else cur["els"]).append(el)
    finish()

    questions, warnings = [], []
    for i, c in enumerate(out, 1):
        qtype, key = parse_answer(c["raw"] or "", c["hint"])
        body = html_of(c["els"])
        body = LEAD_NUM.sub("<p>", body, count=1)
        plain = " ".join(LH.fromstring("<div>" + body + "</div>").text_content().split())
        label_style = "alpha" if re.search(r"\(\s*A\s*\)", plain) else (
            "num" if re.search(r"\(\s*1\s*\)", plain) else "alpha")
        marks, neg = DEFAULT_MARKS[qtype]
        questions.append({
            "position": i, "orig_no": c["orig"], "section": c["section"][:300], "qtype": qtype, "answer": key,
            "label_style": label_style, "html": body, "context": html_of(c["ctx"]),
            "explanation": html_of(c["expl"]), "marks": marks, "negative": neg})
        if not key:
            warnings.append(f"Q{i}: no answer key detected - please enter it.")
    if not questions:
        warnings.append("No questions were detected. Questions must start with a number like '1.' or '12.'.")
    return questions, warnings


def create_exam(file_storage, title: str, duration: int) -> tuple[Exam, list[str]]:
    with tempfile.TemporaryDirectory(prefix="exam_") as tmp:
        job = Path(tmp)
        docx = job / "source.docx"
        file_storage.save(docx)
        validate_docx(docx)
        html = docx_to_html(docx, job)
        questions, warnings = parse_questions(html)
        if not questions:
            raise ValueError(warnings[0])

        code = secrets.token_hex(3).upper()
        while Exam.query.filter_by(code=code).first():
            code = secrets.token_hex(3).upper()
        exam = Exam(title=title, code=code, duration=duration)
        db.session.add(exam)
        db.session.flush()

        names = set()
        media = job / "media"
        if media.exists():
            for p in sorted(media.rglob("*")):
                if p.is_file():
                    names.add(p.name)
                    db.session.add(ExamAsset(
                        exam_id=exam.id, name=p.name, data=p.read_bytes(),
                        mimetype=mimetypes.guess_type(p.name)[0] or "application/octet-stream"))

        def fix_imgs(fragment: str) -> str:
            if not fragment:
                return fragment
            root = LH.fragment_fromstring(fragment, create_parent="div")
            for img in root.iter("img"):
                base = (img.get("src") or "").replace("\\", "/").rsplit("/", 1)[-1]
                if base in names:
                    img.set("src", url_for("exam.asset", exam_id=exam.id, name=base))
                    img.set("alt", "figure")
            return "".join(LH.tostring(c, encoding="unicode") for c in root.iterchildren()) \
                if len(root) else fragment

        for q in questions:
            q["html"] = fix_imgs(q["html"])
            q["context"] = fix_imgs(q["context"])
            q["explanation"] = fix_imgs(q["explanation"])
            exam.questions.append(ExamQuestion(**q))
        db.session.commit()
        return exam, warnings


# ----------------------------------------------------------------- grading
def grade(q: ExamQuestion, given) -> tuple[float, str]:
    """-> (marks delta, status) where status in correct/partial/wrong/skipped."""
    given = (given or "").strip() if isinstance(given, str) else ""
    if not given or not q.answer:
        return 0.0, "skipped"
    if q.qtype == "integer":
        try:
            ok = abs(float(given) - float(q.answer)) <= 0.005 + 1e-9
        except ValueError:
            ok = False
        return (q.marks, "correct") if ok else (-q.negative, "wrong")
    chosen, key = set(given.split(",")), set(q.answer.split(","))
    if chosen == key:
        return q.marks, "correct"
    if q.qtype == "multiple" and chosen < key:
        return round(q.marks * len(chosen) / len(key), 2), "partial"
    return -q.negative, "wrong"


def finalize(att: Attempt) -> None:
    if att.submitted:
        return
    details, score = [], 0.0
    for q in att.exam.questions:
        given = att.answers.get(str(q.id), "")
        delta, status = grade(q, given)
        score += delta
        details.append({"qid": q.id, "given": given, "delta": delta, "status": status})
    att.score, att.total, att.details = round(score, 2), att.exam.total_marks, details
    att.submitted = min(now(), att.deadline)
    db.session.commit()


# -------------------------------------------------------------- admin auth
@bp.record_once
def _setup(state):
    app = state.app
    if not os.environ.get("EXAMINER_PASSWORD"):
        f = Path(app.instance_path) / "admin_password"
        if not f.exists():
            f.write_text(secrets.token_urlsafe(9))
            print(f"\n*** Examiner password generated: {f.read_text()}  (stored in {f}) ***\n")


def admin_password() -> str:
    if os.environ.get("EXAMINER_PASSWORD"):
        return os.environ["EXAMINER_PASSWORD"]
    return (Path(current_app.instance_path) / "admin_password").read_text().strip()


def admin_required(fn):
    @wraps(fn)
    def wrapper(*a, **kw):
        if not session.get("admin"):
            return redirect(url_for("exam.admin_login", next=request.path))
        return fn(*a, **kw)
    return wrapper


@bp.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    if request.method == "POST":
        if secrets.compare_digest(request.form.get("password", ""), admin_password()):
            session["admin"] = True
            nxt = request.args.get("next", "")
            return redirect(nxt if nxt.startswith("/admin") else url_for("exam.admin"))
        flash("Wrong password.", "error")
    return render_template("admin_login.html")


@bp.get("/admin/logout")
def admin_logout():
    session.pop("admin", None)
    return redirect(url_for("exam.landing"))


# ------------------------------------------------------------- admin pages
@bp.get("/")
def landing():
    return render_template("landing.html")


@bp.get("/admin")
@admin_required
def admin():
    exams = Exam.query.order_by(Exam.created.desc()).all()
    counts = {e.id: sum(1 for a in e.attempts if a.submitted) for e in exams}
    return render_template("admin.html", exams=exams, counts=counts,
                           pandoc_ok=pandoc_available())


@bp.post("/admin/upload")
@admin_required
def admin_upload():
    f = request.files.get("file")
    title = (request.form.get("title") or "").strip() or "Untitled exam"
    try:
        duration = max(1, min(600, int(request.form.get("duration", 60))))
    except ValueError:
        duration = 60
    if not pandoc_available():
        flash("Pandoc is not installed on the server.", "error")
        return redirect(url_for("exam.admin"))
    if not f or not f.filename.lower().endswith(".docx"):
        flash("Please choose a .docx file.", "error")
        return redirect(url_for("exam.admin"))
    try:
        exam, warnings = create_exam(f, title[:200], duration)
    except (ValueError, RuntimeError) as e:
        db.session.rollback()
        flash(str(e), "error")
        return redirect(url_for("exam.admin"))
    for w in warnings[:15]:
        flash(w, "warn")
    flash(f"Parsed {len(exam.questions)} questions. Review them below, then publish.", "ok")
    return redirect(url_for("exam.admin_exam", exam_id=exam.id))


def get_exam(exam_id: int) -> Exam:
    exam = db.session.get(Exam, exam_id)
    if exam is None:
        abort(404)
    return exam


@bp.get("/admin/exam/<int:exam_id>")
@admin_required
def admin_exam(exam_id: int):
    exam = get_exam(exam_id)
    return render_template("admin_exam.html", exam=exam, TYPE_LABELS=TYPE_LABELS,
                           join_url=url_for("exam.join", code=exam.code, _external=True))


def _num(v, default=0.0) -> float:
    try:
        return max(0.0, float(v))
    except (TypeError, ValueError):
        return default


@bp.post("/admin/exam/<int:exam_id>/save")
@admin_required
def admin_save(exam_id: int):
    exam, f = get_exam(exam_id), request.form
    exam.title = (f.get("title") or exam.title).strip()[:200]
    try:
        exam.duration = max(1, min(600, int(f.get("duration", exam.duration))))
    except ValueError:
        pass
    exam.show_answers = f.get("show_answers") == "on"
    for q in list(exam.questions):
        if f.get(f"del_{q.id}") == "on":
            db.session.delete(q)
            continue
        t = f.get(f"type_{q.id}")
        q.qtype = t if t in TYPE_LABELS else q.qtype
        q.answer = normalize_key(q.qtype, f.get(f"ans_{q.id}", ""))
        q.marks = _num(f.get(f"marks_{q.id}"), q.marks)
        q.negative = _num(f.get(f"neg_{q.id}"), q.negative)
    db.session.flush()
    for i, q in enumerate(ExamQuestion.query.filter_by(exam_id=exam.id)
                          .order_by(ExamQuestion.position).all(), 1):
        q.position = i
    action = f.get("action", "save")
    if action == "publish":
        missing = [q.position for q in exam.questions if not q.answer]
        if missing or not exam.questions:
            db.session.commit()
            flash(("Cannot publish: missing answer key for Q" + ", Q".join(map(str, missing[:20])))
                  if missing else "Cannot publish: no questions.", "error")
            return redirect(url_for("exam.admin_exam", exam_id=exam.id))
        exam.published = True
        flash("Published. Share the link or code with students.", "ok")
    elif action == "unpublish":
        exam.published = False
        flash("Exam unpublished.", "ok")
    else:
        flash("Saved.", "ok")
    db.session.commit()
    return redirect(url_for("exam.admin_exam", exam_id=exam.id))


@bp.post("/admin/exam/<int:exam_id>/delete")
@admin_required
def admin_delete(exam_id: int):
    exam = get_exam(exam_id)
    ExamAsset.query.filter_by(exam_id=exam.id).delete()
    db.session.delete(exam)
    db.session.commit()
    flash("Exam deleted.", "ok")
    return redirect(url_for("exam.admin"))


@bp.get("/admin/exam/<int:exam_id>/results")
@admin_required
def admin_results(exam_id: int):
    exam = get_exam(exam_id)
    done = sorted((a for a in exam.attempts if a.submitted), key=lambda a: -(a.score or 0))
    return render_template("admin_results.html", exam=exam, attempts=done,
                           pending=sum(1 for a in exam.attempts if not a.submitted))


@bp.get("/admin/exam/<int:exam_id>/results.csv")
@admin_required
def admin_results_csv(exam_id: int):
    exam = get_exam(exam_id)
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["Name", "Roll", "Score", "Total", "Correct", "Wrong", "Skipped", "Started (UTC)", "Minutes taken"])
    for a in exam.attempts:
        if not a.submitted:
            continue
        st = [d["status"] for d in a.details]
        w.writerow([a.name, a.roll, a.score, a.total, st.count("correct") + st.count("partial"),
                    st.count("wrong"), st.count("skipped"), a.started.isoformat(sep=" ", timespec="seconds"),
                    round((a.submitted - a.started).total_seconds() / 60, 1)])
    name = re.sub(r"[^\w.-]+", "_", exam.title)[:60] or "results"
    return Response(buf.getvalue(), mimetype="text/csv",
                    headers={"Content-Disposition": f'attachment; filename="{name}_results.csv"'})


@bp.get("/exam/<int:exam_id>/media/<path:name>")
def asset(exam_id: int, name: str):
    exam = get_exam(exam_id)
    if not exam.published and not session.get("admin"):
        abort(404)
    a = ExamAsset.query.filter_by(exam_id=exam_id, name=name).first_or_404()
    return Response(a.data, mimetype=a.mimetype,
                    headers={"Cache-Control": "private, max-age=3600"})


# ----------------------------------------------------------- student flow
def published_exam(code: str) -> Exam:
    exam = Exam.query.filter_by(code=code.strip().upper()).first()
    if exam is None or not exam.published:
        abort(404)
    return exam


@bp.get("/join")
def join_redirect():
    code = (request.args.get("code") or "").strip().upper()
    return redirect(url_for("exam.join", code=code or "-"))


@bp.get("/e/<code>")
def join(code: str):
    return render_template("join.html", exam=published_exam(code))


@bp.post("/e/<code>/start")
def start(code: str):
    exam = published_exam(code)
    name = (request.form.get("name") or "").strip()[:120]
    roll = (request.form.get("roll") or "").strip()[:60]
    if not name or not roll:
        flash("Please enter your name and roll number.", "error")
        return redirect(url_for("exam.join", code=exam.code))
    att = Attempt.query.filter(Attempt.exam_id == exam.id,
                               db.func.lower(Attempt.roll) == roll.lower()).first()
    if att is None:
        att = Attempt(exam_id=exam.id, token=uuid.uuid4().hex, name=name, roll=roll, answers={})
        db.session.add(att)
        db.session.commit()
    return redirect(url_for("exam.attempt", token=att.token))


def get_attempt(token: str) -> Attempt:
    att = Attempt.query.filter_by(token=token).first()
    if att is None:
        abort(404)
    return att


@bp.get("/attempt/<token>")
def attempt(token: str):
    att = get_attempt(token)
    if not att.submitted and att.remaining <= 0:
        finalize(att)
    if att.submitted:
        return redirect(url_for("exam.result", token=token))
    return render_template("attempt.html", att=att, exam=att.exam, TYPE_LABELS=TYPE_LABELS,
                           LETTERS=LETTERS)


@bp.post("/api/attempt/<token>/save")
def save_answer(token: str):
    att = get_attempt(token)
    data = request.get_json(silent=True) or {}
    if att.submitted or now() > att.deadline + timedelta(seconds=GRACE_SECONDS):
        return jsonify(error="closed"), 409
    qid, value = str(data.get("qid")), str(data.get("value", ""))[:40]
    q = next((q for q in att.exam.questions if str(q.id) == qid), None)
    if q is None:
        return jsonify(error="bad question"), 400
    if q.qtype == "integer":
        value = value.strip() if re.fullmatch(r"-?\d*\.?\d*", value.strip()) else ""
    else:
        value = ",".join(sorted({v for v in value.split(",") if v in LETTERS}))
    answers = dict(att.answers)                      # new object so the JSON change is detected
    if value:
        answers[qid] = value
    else:
        answers.pop(qid, None)
    att.answers = answers
    db.session.commit()
    return jsonify(ok=True, remaining=att.remaining)


@bp.post("/api/attempt/<token>/submit")
def submit(token: str):
    att = get_attempt(token)
    finalize(att)
    return jsonify(redirect=url_for("exam.result", token=token))


@bp.get("/result/<token>")
def result(token: str):
    att = get_attempt(token)
    if not att.submitted:
        return redirect(url_for("exam.attempt", token=token))
    by_id = {d["qid"]: d for d in att.details}
    rows = [(q, by_id.get(q.id)) for q in att.exam.questions]
    stat = [d["status"] for d in att.details]
    return render_template("result.html", att=att, exam=att.exam, rows=rows,
                           correct=stat.count("correct") + stat.count("partial"),
                           wrong=stat.count("wrong"), skipped=stat.count("skipped"))
