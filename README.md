# Online Exams from Word papers  (+ DOCX → LaTeX converter)

**Examiner** uploads a `.docx` question paper → questions, types, answer keys, equations and images are
detected automatically → examiner reviews/edits → publishes → **students** join with a code, take a timed
exam, and are graded automatically.

## Run
```bash
sudo apt install pandoc            # or: brew install pandoc   (2.11+; newer is better)
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
export EXAMINER_PASSWORD='choose-a-password'     # otherwise one is generated in instance/admin_password
python app.py                      # http://localhost:5000
# production: gunicorn -w 2 -b 0.0.0.0:8000 app:app   (also set SECRET_KEY)
```
Pages: `/` home · `/admin` examiner · `/e/<CODE>` student entry · `/converter` DOCX→LaTeX tool.
Everything (including MathJax for equations) is bundled; no internet/CDN is needed at exam time.

## How the paper is read
| In your .docx | Becomes |
|---|---|
| Line starting `1.` / `12)` | a new question (the number is removed; original number kept as "Passage ref") |
| ALL-CAPS line (e.g. `MULTIPLE CORRECT TYPE QUESTIONS`) | section; also sets the default type (Single / Multiple / Integer-Numerical) |
| `(A) … (B) …` or `(1) … (2) …` | options; students pick A–D (shown 1–4 if the paper uses numbers) |
| `Ans. (C)`, `Ans. (A,C,D)`, `Ans. 9.30` | answer key; multi-letter keys make it a multiple-correct question |
| `Sol. …` | explanation, shown in the student's review |
| Text before a group of questions (comprehension) | shared passage shown above each of its questions |
| Equations, tables, images | kept (equations via MathJax, images stored in the database) |

The examiner can fix type, answer, marks and negative marking per question, and delete questions,
before publishing. Publishing is blocked while any answer key is missing.
Defaults: single +4/−1, multiple +4/−2 with partial credit (only if no wrong option is ticked), numerical +4/0 (±0.005).

## Student experience
Name + roll number → timed exam (server-enforced deadline, autosave on every answer, question palette,
resume by roll number if the page is closed) → submit or auto-submit at time-up → score and (optionally) full review.

## Storage: text and images in separate database files
| File (in `instance/`) | Contents |
|---|---|
| `docx2latex.db` | exams, question text/HTML, answer keys, attempts, scores (+ converter history) |
| `media.db` | **only** the exam images (`exam_asset`: name, mimetype, bytes) |

Questions reference images by URL (`/exam/<id>/media/<name>`); the app serves them from `media.db`.
SQLite cannot enforce foreign keys across files, so deleting an exam removes its images explicitly.
Override locations with `DATABASE_URL` / `MEDIA_DATABASE_URL` (e.g. PostgreSQL). Back up both files together.
Changing the schema needs a new DB (no migrations): delete `instance/*.db` if you upgrade from an earlier copy.

## Tests
```bash
python test_exam.py your_paper.docx     # full exam flow: parse, publish, attempt, grading, timer, privacy, both DBs
python test_app.py  your_paper.docx     # LaTeX converter
```

## Known limits
- Answer keys must be A–D / 1–4 or a number; other formats need manual entry in the review screen.
- Students are identified by name + roll only (no accounts); anyone with the code can attempt once per roll number.
- Images that are not PNG/JPEG/GIF/SVG (e.g. EMF/WMF) won't display in browsers.
- Parsing relies on the paper's layout; always skim the review screen (Preview each question) before publishing.
