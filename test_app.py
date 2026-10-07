"""Run: python test_app.py path/to/sample.docx   (needs pandoc). Uses a throwaway DB."""
import io, os, sys, tempfile, zipfile

tmp = tempfile.mkdtemp()
os.environ["DATABASE_URL"] = f"sqlite:///{tmp}/test.db"
from app import app, db, Conversion, Image

path = sys.argv[1]
a, b = app.test_client(), app.test_client()      # two separate browsers (cookie jars)

def upload(c, data, name="sample.docx"):
    return c.post("/api/convert", data={"file": (data, name)}, content_type="multipart/form-data")

# convert + persist
r = upload(a, open(path, "rb")); d = r.get_json(); assert r.status_code == 200, d
print("stats:", d["stats"], "warnings:", d["warnings"])
with app.app_context():
    n_img = Image.query.count(); print("images stored in DB:", n_img)
    assert Conversion.query.count() == 1 and n_img > 0

# output quality fixes
t = d["latex"]
assert "\\frac{" not in t and "\\dfrac{" in t, "inline fractions"
assert "hfill\\break" not in t, "junk line breaks"
assert "\\section*{SINGLE CORRECT TYPE QUESTIONS}" in t, "caps headings"
assert "\\raisebox{-\\height}{\\includegraphics" in t, "table images top-aligned"
off = upload(a, open(path, "rb")).get_json()  # defaults on; now check the toggles off
r2 = a.post("/api/convert", data={"file": (open(path, "rb"), "s.docx"), "bigfrac": "false",
            "headings": "false"}, content_type="multipart/form-data").get_json()
assert "\\dfrac{" not in r2["latex"] and "\\section*{" not in r2["latex"], "toggles"
for x in (off, r2): a.delete(f"/api/conversions/{x['id']}")

# history, reopen, downloads (built from DB, not disk)
h = a.get("/api/history").get_json(); assert len(h) == 1 and "latex" not in h[0]
assert a.get(f"/api/conversions/{d['id']}").get_json()["latex"] == d["latex"]
z = a.get(f"/download/{d['id']}/zip"); assert z.status_code == 200
names = zipfile.ZipFile(io.BytesIO(z.data)).namelist()
assert "main.tex" in names and any(n.startswith("media/") for n in names); print("zip files:", len(names))
assert a.get(f"/download/{d['id']}/tex").data.decode() == d["latex"]

# privacy: another browser sees nothing and cannot fetch/delete it
assert b.get("/api/history").get_json() == []
assert b.get(f"/api/conversions/{d['id']}").status_code == 404
assert b.get(f"/download/{d['id']}/zip").status_code == 404
assert b.delete(f"/api/conversions/{d['id']}").status_code == 404

# validation
assert upload(a, io.BytesIO(b"nope")).status_code == 400
assert upload(a, io.BytesIO(b"x"), "x.txt").status_code == 400
assert a.get("/download/1/evil").status_code == 404

# delete cascades to images
assert a.delete(f"/api/conversions/{d['id']}").status_code == 200
with app.app_context():
    assert Conversion.query.count() == 0 and Image.query.count() == 0
assert a.get("/api/history").get_json() == []
print("all checks passed")
