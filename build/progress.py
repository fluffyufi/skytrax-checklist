"""Renders the live progress page from progress.json -> scratchpad/progress.html"""
import html
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = "/tmp/claude-0/-home-user-skytrax-checklist/c0ab9c35-2f2c-5eff-909f-a4943aca1b4a/scratchpad/progress.html"
st = json.load(open(os.path.join(HERE, "progress.json")))

PILL = {"pass": "Passed", "fail": "Returned for fix", "run": "In progress", "wait": "Queued"}


def e(s):
    return html.escape(str(s))


rows = []
for p in st["pieces"]:
    rows.append(f"""<tr><td class="pc">{e(p['name'])}</td><td><span class="pill {p['state']}">{PILL[p['state']]}</span></td>
<td class="num">{e(p.get('round', ''))}</td><td>{e(p.get('note', ''))}</td></tr>""")
log = "".join(f"<li><time>{e(t)}</time><span>{e(m)}</span></li>" for t, m in reversed(st["log"]))
gaps = "".join(f"<li><b>{e(g['piece'])}</b> · round {e(g['round'])}: {e(g['gap'])}</li>" for g in reversed(st.get("gaps", [])))
done = sum(1 for p in st["pieces"] if p["state"] == "pass")
page = f"""<title>Skytrax Catering Build</title>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;600&family=IBM+Plex+Mono:wght@400;600&display=swap">
<style>
:root{{--bg:#f3f5f8;--panel:#fff;--ink:#14213a;--mute:#5b6679;--line:#d9dfe8;--navy:#1f3864;
--ok:#1d7a4a;--okbg:#e2f3ea;--bad:#b3261e;--badbg:#fbe4e2;--run:#8a5a00;--runbg:#fdf0d5;--wait:#5b6679;--waitbg:#eceff3}}
@media (prefers-color-scheme:dark){{:root:not([data-theme="light"]){{color-scheme:dark;--bg:#0f1522;--panel:#172033;--ink:#e6ebf4;--mute:#9aa6ba;--line:#2a3550;--navy:#9db5e8;
--ok:#6fd39e;--okbg:#153626;--bad:#ff9d94;--badbg:#3d1714;--run:#f2c46b;--runbg:#3a2c0f;--wait:#9aa6ba;--waitbg:#222c40}}}}
:root[data-theme="dark"]{{color-scheme:dark;--bg:#0f1522;--panel:#172033;--ink:#e6ebf4;--mute:#9aa6ba;--line:#2a3550;--navy:#9db5e8;
--ok:#6fd39e;--okbg:#153626;--bad:#ff9d94;--badbg:#3d1714;--run:#f2c46b;--runbg:#3a2c0f;--wait:#9aa6ba;--waitbg:#222c40}}
body{{background:var(--bg);color:var(--ink);font:15px/1.5 "IBM Plex Sans",system-ui,sans-serif;padding:24px 16px}}
main{{max-width:980px;margin:0 auto;display:grid;gap:20px}}
h1{{font-size:24px;margin:0;color:var(--navy);text-wrap:balance}} h2{{font-size:13px;letter-spacing:.08em;text-transform:uppercase;color:var(--mute);margin:0 0 8px}}
.sub{{color:var(--mute);margin:4px 0 0}}
.meter{{display:flex;gap:16px;flex-wrap:wrap;font-family:"IBM Plex Mono",monospace;font-size:13px;color:var(--mute)}} .meter b{{color:var(--ink);font-size:20px;display:block}}
section{{background:var(--panel);border:1px solid var(--line);border-radius:6px;padding:16px}}
.tw{{overflow-x:auto}} table{{border-collapse:collapse;width:100%;min-width:620px}} th,td{{text-align:left;padding:8px 10px;border-bottom:1px solid var(--line);vertical-align:top}}
th{{font-size:12px;color:var(--mute);text-transform:uppercase;letter-spacing:.06em}} .pc{{font-weight:600}} .num{{font-variant-numeric:tabular-nums;font-family:"IBM Plex Mono",monospace}}
.pill{{display:inline-block;padding:2px 9px;border-radius:99px;font-size:12px;font-weight:600;white-space:nowrap}}
.pass{{color:var(--ok);background:var(--okbg)}} .fail{{color:var(--bad);background:var(--badbg)}} .run{{color:var(--run);background:var(--runbg)}} .wait{{color:var(--wait);background:var(--waitbg)}}
ul{{margin:0;padding-left:18px}} li{{margin:4px 0}} .log{{list-style:none;padding:0}} .log li{{display:grid;grid-template-columns:120px 1fr;gap:10px}} time{{font-family:"IBM Plex Mono",monospace;color:var(--mute);font-size:13px}}
@media (max-width:520px){{.log li{{grid-template-columns:1fr}}}}
</style>
<main>
<header><h1>MAGCS Skytrax 2026 catering readiness workbook</h1>
<p class="sub">Builder / critic loop status. Each piece passes only when a fresh critic finds no blocking gap.</p></header>
<div class="meter"><div><b>{done}/{len(st['pieces'])}</b>pieces passed</div><div><b>22</b>flight legs</div><div><b>{e(st.get('checks', '—'))}</b>check lines</div><div><b>{e(st['updated'])}</b>last update (MYT)</div></div>
<section><h2>Pieces</h2><div class="tw"><table><thead><tr><th>Piece</th><th>State</th><th>Round</th><th>Latest</th></tr></thead><tbody>{''.join(rows)}</tbody></table></div></section>
<section><h2>Biggest gap returned per critic round</h2><ul>{gaps or '<li>None yet.</li>'}</ul></section>
<section><h2>Log</h2><ul class="log">{log}</ul></section>
</main>"""
open(OUT, "w").write(page)
print(OUT)
