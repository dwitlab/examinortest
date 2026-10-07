"""End-to-end test of the exam platform. Run: python test_exam.py path/to/paper.docx"""
import os, sqlite3, sys, tempfile
tmp = tempfile.mkdtemp()
os.environ.update(DATABASE_URL=f"sqlite:///{tmp}/exam.db", MEDIA_DATABASE_URL=f"sqlite:///{tmp}/media.db",
                  EXAMINER_PASSWORD="secret-pw")
from datetime import timedelta
from app import app, db
from exam import Exam, ExamQuestion, ExamAsset, Attempt, now

DOCX = sys.argv[1]
admin, student, other = app.test_client(), app.test_client(), app.test_client()

# --- examiner auth
assert admin.get("/admin").status_code == 302, "admin must require login"
assert admin.post("/admin/login", data={"password": "wrong"}).status_code == 200
assert admin.post("/admin/login", data={"password": "secret-pw"}).status_code == 302

# --- upload & parse
r = admin.post("/admin/upload", data={"title": "Physics Test", "duration": "30",
               "file": (open(DOCX, "rb"), "paper.docx")}, content_type="multipart/form-data")
assert r.status_code == 302, r.data
with app.app_context():
    ex = Exam.query.one(); eid, code = ex.id, ex.code
    qs = ex.questions
    print(f"parsed {len(qs)} questions, total marks {ex.total_marks}, code {code}")
    assert len(qs) == 17 and all(q.answer for q in qs)
    types = {t: sum(q.qtype == t for q in qs) for t in ("single", "multiple", "integer")}; print(types)
    assert types == {"single": 11, "multiple": 2, "integer": 4}, types
    n_assets = ExamAsset.query.count(); print("images:", n_assets)
    assert n_assets >= 12
    assert all("/exam/%d/media/" % eid in q.html for q in qs if "<img" in q.html), "img src rewritten"
    assert not any("media/image" in q.html and "/exam/" not in q.html for q in qs)
    key = {q.id: (q.qtype, q.answer, q.marks, q.negative) for q in qs}

# --- two separate database files, images only in media.db
con_t, con_m = sqlite3.connect(f"{tmp}/exam.db"), sqlite3.connect(f"{tmp}/media.db")
tt = {r[0] for r in con_t.execute("select name from sqlite_master where type='table'")}
tm = {r[0] for r in con_m.execute("select name from sqlite_master where type='table'")}
print("exam.db tables:", sorted(tt)); print("media.db tables:", sorted(tm))
assert "exam_asset" in tm and "exam_asset" not in tt and "exam_question" in tt and "exam_question" not in tm
assert con_m.execute("select count(*) from exam_asset").fetchone()[0] == n_assets
assert b"<img" not in con_m.execute("select group_concat(name) from exam_asset").fetchone()[0].encode()

# --- cannot publish with a missing key; can after fixing
qid0 = next(iter(key)); form = {"title": "Physics Test", "duration": "30", "show_answers": "on", "action": "publish"}
for q_id, (t, a, m, n) in key.items():
    form.update({f"type_{q_id}": t, f"ans_{q_id}": a, f"marks_{q_id}": str(m), f"neg_{q_id}": str(n)})
bad = dict(form); bad[f"ans_{qid0}"] = ""
admin.post(f"/admin/exam/{eid}/save", data=bad)
with app.app_context(): assert not db.session.get(Exam, eid).published
assert student.get(f"/e/{code}").status_code == 404, "draft exam invisible to students"
admin.post(f"/admin/exam/{eid}/save", data=form)
with app.app_context(): assert db.session.get(Exam, eid).published
assert student.get(f"/e/{code}").status_code == 200

# --- student takes the exam
assert student.post(f"/e/{code}/start", data={"name": "", "roll": ""}).status_code == 302
r = student.post(f"/e/{code}/start", data={"name": "Asha", "roll": "R1"}); tok = r.headers["Location"].split("/")[-1]
assert student.post(f"/e/{code}/start", data={"name": "Asha", "roll": "r1"}).headers["Location"].endswith(tok), "resume by roll"
page = student.get(f"/attempt/{tok}"); assert page.status_code == 200 and b"Question 1 of 17" in page.data
html = page.data.decode()
for q_id, (t, a, m, n) in key.items():
    assert f'data-qid="{q_id}"' in html
# answers are never sent to the student before submit
for q_id, (t, a, m, n) in key.items():
    assert f'value="{a}" checked' not in html

order = list(key.items()); expected = 0.0
def save(q_id, v): return student.post(f"/api/attempt/{tok}/save", json={"qid": q_id, "value": v})
for i, (q_id, (t, a, m, n)) in enumerate(order):
    if i % 5 == 4: continue                               # skip some
    if i % 5 == 3 and t != "integer":                     # answer wrong on purpose
        wrong = "B" if a != "B" else "C"; assert save(q_id, wrong).status_code == 200
        expected -= n if set([wrong]) != set(a.split(",")) else 0
        continue
    if t == "multiple" and i % 5 == 2:                    # partial credit: a strict subset
        sub = a.split(",")[:1]; save(q_id, ",".join(sub)); expected += round(m * len(sub) / len(a.split(",")), 2); continue
    save(q_id, a); expected += m
assert save(order[0][0], "Z").status_code == 200          # junk is sanitised, not stored
assert save(99999, "A").status_code == 400
# integer sanitising
int_q = next(k for k, v in key.items() if v[0] == "integer")
save(int_q, "abc"); 
with app.app_context(): assert str(int_q) not in db.session.query(Attempt).filter_by(token=tok).one().answers
save(int_q, key[int_q][1])

# recompute expected from DB truth to avoid tracking mistakes in the loop above
with app.app_context():
    att = Attempt.query.filter_by(token=tok).one(); import exam as E
    exp = sum(E.grade(q, att.answers.get(str(q.id), ""))[0] for q in att.exam.questions)
assert student.post(f"/api/attempt/{tok}/submit").status_code == 200
with app.app_context():
    att = Attempt.query.filter_by(token=tok).one()
    print("score", att.score, "/", att.total, "recomputed", round(exp, 2))
    assert att.submitted and abs(att.score - round(exp, 2)) < 1e-6 and att.total == 68.0
assert save(order[1][0], "A").status_code == 409, "no edits after submit"
assert student.get(f"/attempt/{tok}").status_code == 302
res = student.get(f"/result/{tok}"); assert res.status_code == 200 and b"Correct answer" in res.data

# --- grading rules
with app.app_context():
    g = E.grade
    mq = ExamQuestion(qtype="multiple", answer="A,C,D", marks=4, negative=2)
    assert g(mq, "A,C,D") == (4, "correct") and g(mq, "A,C")[1] == "partial" and g(mq, "A,B") == (-2, "wrong") and g(mq, "") == (0, "skipped")
    iq = ExamQuestion(qtype="integer", answer="9.30", marks=4, negative=0)
    assert g(iq, "9.3")[1] == "correct" and g(iq, "9.4")[1] == "wrong"
    sq = ExamQuestion(qtype="single", answer="C", marks=4, negative=1)
    assert g(sq, "C")[0] == 4 and g(sq, "A")[0] == -1

# --- timer enforcement: expired attempt is auto-graded, saves rejected
tok2 = other.post(f"/e/{code}/start", data={"name": "Ravi", "roll": "R2"}).headers["Location"].split("/")[-1]
with app.app_context():
    a2 = Attempt.query.filter_by(token=tok2).one(); a2.started = now() - timedelta(minutes=31); db.session.commit()
assert other.post(f"/api/attempt/{tok2}/save", json={"qid": order[0][0], "value": "A"}).status_code == 409
assert other.get(f"/attempt/{tok2}").status_code == 302
with app.app_context(): a2 = Attempt.query.filter_by(token=tok2).one(); assert a2.submitted and a2.score == 0

# --- media access & admin protection
img = next(a for a in [con_m.execute("select name from exam_asset limit 1").fetchone()])[0]
assert student.get(f"/exam/{eid}/media/{img}").status_code == 200
assert student.get("/admin/exam/1").status_code == 302 and student.get(f"/admin/exam/{eid}/results.csv").status_code == 302
assert student.get(f"/exam/{eid}/media/../x").status_code in (404, 308)

# --- results & csv
r = admin.get(f"/admin/exam/{eid}/results"); assert b"Asha" in r.data and b"Ravi" in r.data
csv_ = admin.get(f"/admin/exam/{eid}/results.csv"); assert csv_.status_code == 200 and b"Asha,R1" in csv_.data

# --- unpublish hides, delete cleans BOTH databases
admin.post(f"/admin/exam/{eid}/save", data={**form, "action": "unpublish"}); assert student.get(f"/e/{code}").status_code == 404
admin.post(f"/admin/exam/{eid}/delete")
assert con_m.execute("select count(*) from exam_asset").fetchone()[0] == 0
assert sqlite3.connect(f"{tmp}/exam.db").execute("select count(*) from exam_question").fetchone()[0] == 0
print("ALL EXAM CHECKS PASSED")
