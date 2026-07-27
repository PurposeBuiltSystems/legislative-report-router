#!/usr/bin/env python3
"""Render docs/*.md to styled HTML siblings + a docs index.
Run after editing any doc: python3 scripts/render_docs.py"""
import glob, html, os, re

STYLE = """<style>
body { font-family: -apple-system, "Segoe UI", sans-serif; margin: 3rem auto; max-width: 46rem;
  padding: 0 1rem; color: #242424; line-height: 1.6; }
h1, h2, h3 { color: #0f6cbd; } h1 { font-size: 1.6em; } h2 { margin-top: 2rem; }
table { border-collapse: collapse; width: 100%; font-size: 0.93em; margin: 12px 0; }
th, td { border: 1px solid #d1d1d1; padding: 6px 9px; text-align: left; vertical-align: top; }
th { background: #eef4f6; }
code { background: #f3f2f1; padding: 0.1rem 0.35rem; border-radius: 3px; font-size: 0.92em; }
pre { background: #f3f2f1; padding: 10px 12px; border-radius: 6px; overflow-x: auto; }
pre code { background: none; padding: 0; }
blockquote { border-left: 4px solid #0f6cbd; margin: 12px 0; padding: 4px 14px; color: #424242; background: #f8fafc; }
a { color: #0f6cbd; } .crumb { font-size: 0.9em; }
@media print { body { margin: 0.5in; } }
</style>"""

def inline(s):
    s = html.escape(s, quote=False)
    s = re.sub(r"`([^`]+)`", r"<code>\1</code>", s)
    s = re.sub(r"\*\*([^*]+)\*\*", r"<b>\1</b>", s)
    s = re.sub(r"\[([^\]]+)\]\(([^)\s]+)\)", r'<a href="\2">\1</a>', s)
    return s

def render(md, title):
    out, para = [], []
    mode = [None]  # ul/ol/table
    def flush_para():
        if para:
            out.append("<p>" + "<br>".join(para) + "</p>")
            del para[:]
    def close(m=None):
        flush_para()
        if mode[0] == "ul": out.append("</ul>")
        if mode[0] == "ol": out.append("</ol>")
        if mode[0] == "table": out.append("</table>")
        mode[0] = m
    in_code = False
    for ln in md.split("\n"):
        if ln.strip().startswith("```"):
            close()
            out.append("<pre><code>" if not in_code else "</code></pre>")
            in_code = not in_code
            continue
        if in_code:
            out.append(html.escape(ln))
            continue
        st = ln.strip()
        if st.startswith("|"):
            cells = [c.strip() for c in st.strip("|").split("|")]
            if all(re.fullmatch(r":?-{2,}:?", c) for c in cells if c):
                continue
            if mode[0] != "table":
                close("table")
                out.append("<table>")
                out.append("<tr>" + "".join("<th>" + inline(c) + "</th>" for c in cells) + "</tr>")
            else:
                out.append("<tr>" + "".join("<td>" + inline(c) + "</td>" for c in cells) + "</tr>")
            continue
        m = re.match(r"^(#{1,3})\s+(.*)", st)
        if m:
            close()
            lvl = len(m.group(1))
            txt = m.group(2)
            anchor = re.sub(r"[^a-z0-9]+", "-", txt.lower()).strip("-")
            out.append('<h%d id="%s">%s</h%d>' % (lvl, anchor, inline(txt), lvl))
            continue
        if st.startswith("- "):
            if mode[0] != "ul":
                close("ul")
                out.append("<ul>")
            out.append("<li>" + inline(st[2:]) + "</li>")
            continue
        m = re.match(r"^(\d+)\.\s+(.*)", st)
        if m:
            if mode[0] != "ol":
                close("ol")
                out.append('<ol start="%s">' % m.group(1))
            out.append("<li>" + inline(m.group(2)) + "</li>")
            continue
        if st.startswith("> "):
            close()
            out.append("<blockquote>" + inline(st[2:]) + "</blockquote>")
            continue
        if not st:
            close()
            continue
        if mode[0] in ("ul", "ol") and ln.startswith("  ") and out and out[-1].endswith("</li>"):
            out[-1] = out[-1][:-5] + " " + inline(st) + "</li>"
            continue
        if mode[0]:
            close()
        para.append(inline(st))
    close()
    return ('<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">'
            '<meta name="viewport" content="width=device-width, initial-scale=1.0">'
            "<title>" + html.escape(title) + " — Legislative Report Router</title>" + STYLE +
            '</head><body><p class="crumb"><a href="index.html">← Documentation</a></p>' +
            "\n".join(out) + "</body></html>")

DESC = {
    "admin-guide": "List schemas, Teams tab setup, imports, per-state config — the administrator's manual",
    "permissions": "Every Graph permission, why it's needed, and the Sites.Selected hardening option",
    "parser-rules": "Exactly how bill reports are parsed — deterministic rules, no AI",
    "government-cloud": "GCC / GCC High / DoD endpoint configuration",
    "copilot-agent": "Build the Bill Summarizer agent (included with M365 Copilot licenses)",
    "agent-builder-business-case": "The case for enabling Copilot Agent Builder, written for IT",
}

index_rows = []
for path in sorted(glob.glob("docs/*.md")):
    name = os.path.splitext(os.path.basename(path))[0]
    md = open(path).read()
    m = re.match(r"^#\s+(.*)", md.strip())
    title = m.group(1) if m else name
    open("docs/" + name + ".html", "w").write(render(md, title))
    index_rows.append((name, title, DESC.get(name, "")))
    print("rendered", name + ".html")

idx = ['<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">',
       '<meta name="viewport" content="width=device-width, initial-scale=1.0">',
       "<title>Documentation — Legislative Report Router</title>", STYLE, "</head><body>",
       "<h1>Documentation</h1>",
       '<p><a href="../quickstart.html"><b>New coordinator? Start with the Quick Start →</b></a></p>']
for name, title, desc in index_rows:
    idx.append('<p><a href="%s.html"><b>%s</b></a><br><span style="color:#616161">%s</span></p>' % (name, html.escape(title), html.escape(desc)))
idx.append("</body></html>")
open("docs/index.html", "w").write("\n".join(idx))
print("rendered index.html")
