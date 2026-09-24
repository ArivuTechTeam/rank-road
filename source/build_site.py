"""Rebuild everything: Excel timetables (../excel) and the four web apps (../site).

    cd source
    pip install openpyxl pillow
    python build_site.py

Edit syllabus/hours/dates in build.py (Foundation) or inter_data.py + gen_inter.py (Inter),
then run this script and redeploy the ../site folder.
"""
import os, re, subprocess, sys
HERE = os.path.dirname(os.path.abspath(__file__)); os.chdir(HERE)
os.makedirs("../excel", exist_ok=True)
subprocess.run([sys.executable, "build.py"], check=True)       # -> ../excel/CA_Foundation...xlsx + plan.json
subprocess.run([sys.executable, "gen_inter.py"], check=True)   # -> ../excel/CA_Inter...xlsx + plan_g1/g2/both.json

HEAD = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="theme-color" content="#1F4E3D">
<meta name="description" content="{desc}">
<link rel="manifest" href="../manifest.webmanifest">
<link rel="icon" href="../assets/icon-192.png">
<link rel="apple-touch-icon" href="../assets/icon-192.png">
<style>
:root{{color-scheme:light;padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}}
body{{margin:0;font:14px/1.5 system-ui,-apple-system,"Segoe UI",sans-serif}}
img{{max-width:100%}}[hidden]{{display:none!important}}
</style>
"""
TAIL = """
<script>if("serviceWorker" in navigator){addEventListener("load",()=>navigator.serviceWorker.register("../sw.js").catch(()=>{}));}</script>
</body>
</html>
"""
def page(tpl, plan_json, title, folder, desc):
    body = open(tpl, encoding="utf-8").read().replace("__TITLE__", title).replace("__PLAN__", open(plan_json, encoding="utf-8").read())
    m = re.match(r"\s*(<title>.*?</title>)", body, re.S); t = m.group(1); body = body[m.end():]
    i = body.index('<div class="wrap">')
    os.makedirs(f"../site/{folder}", exist_ok=True)
    open(f"../site/{folder}/index.html", "w", encoding="utf-8").write(HEAD.format(desc=desc) + t + "\n" + body[:i] + "</head>\n<body>\n" + body[i:] + TAIL)

page("foundation_template.html", "plan.json", "Rank Road Jan 2027", "foundation", "CA Foundation January 2027 day-wise study tracker")
page("inter_template.html", "plan_g1.json", "Rank Road Inter Group I", "inter-group1", "CA Intermediate Group I January 2027 day-wise study tracker")
page("inter_template.html", "plan_g2.json", "Rank Road Inter Group II", "inter-group2", "CA Intermediate Group II January 2027 day-wise study tracker")
page("inter_template.html", "plan_both.json", "Rank Road Inter Both Groups", "inter-both", "CA Intermediate both groups January 2027 day-wise study tracker")
print("Site rebuilt in ../site — remember to bump VERSION in ../site/sw.js before redeploying.")
