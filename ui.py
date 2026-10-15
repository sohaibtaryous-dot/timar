# -*- coding: utf-8 -*-
"""Habillage visuel de l'application : thème, en-tête 3D animé, cartes KPI 3D, graphiques."""

from __future__ import annotations

import html


import altair as alt
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# Palette de l'interface (thème sombre)
BG = "#070D1A"
SURFACE = "#0F1729"
INK = "#E8EEF8"
INK_2 = "#9AA8C2"
INK_3 = "#64718C"
ACCENT = "#3B82F6"
ACCENT_2 = "#22D3EE"
GOLD = "#F5B83D"
# Couleurs de statut réservées (toujours accompagnées d'un libellé + icône)
GOOD, WARNING, CRITICAL = "#0CA30C", "#FAB219", "#D03B3B"
SERIES_1 = "#3987E5"  # série unique des graphiques

FONTS = ("https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800"
         "&family=Space+Grotesk:wght@500;600;700&display=swap")


# =============================================================================
# CSS GLOBAL
# =============================================================================

GLOBAL_CSS = f"""
<style>
@import url('{FONTS}');

:root {{
  --bg: {BG}; --surface: {SURFACE}; --ink: {INK}; --ink2: {INK_2}; --ink3: {INK_3};
  --accent: {ACCENT}; --accent2: {ACCENT_2}; --gold: {GOLD};
  --glass: rgba(17, 27, 48, .62); --glass-brd: rgba(120, 160, 255, .16);
}}
html, body, [class*="css"], .stApp, .stMarkdown, button, input, textarea, select {{
  font-family: 'Inter', system-ui, sans-serif !important;
}}
h1, h2, h3, h4 {{ font-family: 'Space Grotesk', 'Inter', sans-serif !important; letter-spacing: -.01em; }}

/* ---------- Fond animé ---------- */
.stApp {{ background: var(--bg); }}
.stApp::before {{
  content: ""; position: fixed; inset: -20%; z-index: 0; pointer-events: none;
  background:
    radial-gradient(38% 32% at 18% 12%, rgba(59,130,246,.22), transparent 70%),
    radial-gradient(30% 28% at 85% 18%, rgba(34,211,238,.14), transparent 70%),
    radial-gradient(35% 30% at 70% 90%, rgba(245,184,61,.08), transparent 70%);
  animation: drift 22s ease-in-out infinite alternate;
}}
.stApp::after {{
  content: ""; position: fixed; inset: 0; z-index: 0; pointer-events: none; opacity: .35;
  background-image: linear-gradient(rgba(120,160,255,.06) 1px, transparent 1px),
                    linear-gradient(90deg, rgba(120,160,255,.06) 1px, transparent 1px);
  background-size: 44px 44px;
  mask-image: radial-gradient(ellipse at 50% 0%, #000 30%, transparent 75%);
}}
@keyframes drift {{ from {{ transform: translate3d(0,0,0) rotate(0deg); }} to {{ transform: translate3d(4%, 3%, 0) rotate(6deg); }} }}
[data-testid="stAppViewContainer"] > .main, .block-container {{ position: relative; z-index: 1; }}
.block-container {{ padding-top: 1.2rem !important; padding-bottom: 3rem; max-width: 1380px; }}
header[data-testid="stHeader"] {{ background: transparent; }}
#MainMenu, footer, [data-testid="stDecoration"] {{ display: none !important; }}

/* ---------- Barre latérale ---------- */
section[data-testid="stSidebar"] {{
  background: linear-gradient(180deg, rgba(13,22,42,.96), rgba(7,13,26,.98));
  border-right: 1px solid var(--glass-brd);
}}
section[data-testid="stSidebar"] h2, section[data-testid="stSidebar"] h3 {{ font-size: 1rem !important; color: var(--ink); }}
.side-brand {{ display:flex; align-items:center; gap:.7rem; padding:.2rem 0 1rem; border-bottom:1px solid var(--glass-brd); margin-bottom:1rem; }}
.side-brand .logo {{ width:42px; height:42px; border-radius:12px; display:grid; place-items:center; font-weight:800;
  font-family:'Space Grotesk'; color:white; background:linear-gradient(135deg,var(--accent),var(--accent2));
  box-shadow:0 8px 24px rgba(59,130,246,.45); }}
.side-brand b {{ display:block; color:var(--ink); font-size:.98rem; }}
.side-brand span {{ color:var(--ink3); font-size:.75rem; letter-spacing:.08em; text-transform:uppercase; }}
.side-pill {{ display:flex; gap:.5rem; align-items:center; font-size:.82rem; padding:.55rem .75rem; border-radius:10px;
  border:1px solid var(--glass-brd); background:var(--glass); color:var(--ink2); margin:.4rem 0 .8rem; }}
.side-pill i {{ width:8px; height:8px; border-radius:50%; display:inline-block; flex:none; }}

/* ---------- Sélecteur de module (Import / Export) ---------- */
div.st-key-module {{ width:100% !important; }}
div.st-key-module [data-testid="stWidgetLabel"] {{ display:none; }}
div.st-key-module [role="radiogroup"] {{ gap:.5rem; width:100%; display:flex; flex-direction:column; }}
div.st-key-module [role="radiogroup"] > div {{ width:100%; }}
div.st-key-module [data-testid="stRadioOption"], div.st-key-module [role="radiogroup"] > label {{
  width:100%; margin:0; padding:.75rem .9rem; border-radius:12px; border:1px solid var(--glass-brd);
  background:var(--glass); transition:all .25s; cursor:pointer; }}
div.st-key-module [data-testid="stRadioOption"]:hover {{ border-color:rgba(59,130,246,.5); transform:translateX(3px); }}
div.st-key-module [data-testid="stRadioOption"][data-selected="true"], div.st-key-module [role="radiogroup"] > label:has(input:checked) {{
  background:linear-gradient(135deg,rgba(37,99,235,.5),rgba(34,211,238,.22)); border-color:rgba(96,165,250,.75);
  box-shadow:0 10px 26px -12px rgba(59,130,246,.9); }}
div.st-key-module [data-testid="stRadioOption"] p {{ font-weight:700; font-size:.86rem; color:var(--ink); }}
div.st-key-module [data-testid="stRadioOption"] > div > div:first-child:not([data-testid]) {{ display:none; }}
.side-cap {{ font-size:.7rem; letter-spacing:.14em; text-transform:uppercase; color:var(--ink3); margin:.2rem 0 .45rem; }}

/* ---------- Cartes « verre » (conteneurs à clé) ---------- */
div[class*="st-key-card"] {{
  background: var(--glass); border: 1px solid var(--glass-brd); border-radius: 20px;
  padding: 1.1rem 1.2rem 1rem; backdrop-filter: blur(14px); -webkit-backdrop-filter: blur(14px);
  box-shadow: 0 20px 50px -24px rgba(0,0,0,.7), inset 0 1px 0 rgba(255,255,255,.04);
  transition: transform .35s cubic-bezier(.2,.8,.2,1), box-shadow .35s, border-color .35s;
  animation: rise .7s cubic-bezier(.2,.8,.2,1) both;
}}
div[class*="st-key-card_upl"]:hover {{
  transform: perspective(900px) rotateX(2deg) translateY(-4px);
  border-color: rgba(59,130,246,.45); box-shadow: 0 30px 60px -24px rgba(59,130,246,.35);
}}
@keyframes rise {{ from {{ opacity:0; transform: translateY(16px); }} to {{ opacity:1; transform:none; }} }}
.card-title {{ display:flex; gap:.85rem; align-items:center; margin-bottom:.6rem; }}
.card-title .ico {{ width:44px; height:44px; border-radius:14px; display:grid; place-items:center; font-size:1.3rem; flex:none;
  background: linear-gradient(135deg, rgba(59,130,246,.25), rgba(34,211,238,.12)); border:1px solid var(--glass-brd); }}
.card-title b {{ color: var(--ink); font-size: 1.02rem; display:block; font-family:'Space Grotesk'; }}
.card-title span {{ color: var(--ink2); font-size: .84rem; }}
.doc-tags {{ display:flex; flex-wrap:wrap; gap:.35rem; margin:.1rem 0 .5rem; }}
.doc-tags em {{ font-style:normal; font-size:.72rem; color:var(--ink2); padding:.18rem .55rem; border-radius:999px;
  border:1px solid var(--glass-brd); background:rgba(59,130,246,.08); }}

/* Zone de dépôt */
[data-testid="stFileUploaderDropzone"] {{
  background: rgba(7,13,26,.55) !important; border: 1.5px dashed rgba(96,140,230,.45) !important;
  border-radius: 14px !important; transition: all .25s;
}}
[data-testid="stFileUploaderDropzone"]:hover {{ border-color: var(--accent2) !important; background: rgba(34,211,238,.06) !important; }}
[data-testid="stFileUploader"] label {{ display:none; }}

/* ---------- Boutons ---------- */
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {{
  background: linear-gradient(135deg, #2563EB 0%, #3B82F6 45%, #06B6D4 100%) !important;
  background-size: 200% 200% !important; border: 0 !important; color: white !important;
  font-weight: 700 !important; letter-spacing: .02em; border-radius: 14px !important; min-height: 3.2rem;
  box-shadow: 0 12px 30px -10px rgba(59,130,246,.75), inset 0 1px 0 rgba(255,255,255,.25);
  transition: transform .2s, box-shadow .2s, background-position .6s !important;
}}
.stButton > button[kind="primary"]:hover, .stDownloadButton > button[kind="primary"]:hover {{
  transform: translateY(-2px); background-position: 100% 0 !important;
  box-shadow: 0 18px 40px -10px rgba(34,211,238,.7), inset 0 1px 0 rgba(255,255,255,.3);
}}
.stButton > button[kind="primary"]:disabled {{ background: #1A2540 !important; color: var(--ink3) !important; box-shadow:none; }}
.stDownloadButton > button[kind="primary"] {{
  background: linear-gradient(135deg, #059669, #10B981 50%, #22D3EE) !important;
  box-shadow: 0 12px 30px -10px rgba(16,185,129,.7) !important; animation: pulse 2.4s ease-in-out infinite;
}}
@keyframes pulse {{ 0%,100% {{ box-shadow: 0 12px 30px -10px rgba(16,185,129,.7); }} 50% {{ box-shadow: 0 12px 44px -6px rgba(16,185,129,.95); }} }}
.stButton > button[kind="secondary"] {{ border-radius: 12px; border: 1px solid var(--glass-brd); background: var(--glass); }}

/* ---------- Progression ---------- */
.stProgress > div > div > div {{ background: rgba(255,255,255,.06) !important; border-radius: 99px; height: 10px !important; }}
.stProgress > div > div > div > div {{
  background: linear-gradient(90deg, #2563EB, #22D3EE, #2563EB) !important; background-size: 200% 100% !important;
  animation: flow 1.6s linear infinite; border-radius: 99px; box-shadow: 0 0 18px rgba(34,211,238,.6);
}}
@keyframes flow {{ from {{ background-position: 0 0; }} to {{ background-position: 200% 0; }} }}

/* ---------- Onglets (anciens et nouveaux DOM Streamlit) ---------- */
.stTabs [role="tablist"] {{ gap: .4rem; background: var(--glass); padding: .4rem; border-radius: 14px;
  border: 1px solid var(--glass-brd); width: max-content; max-width: 100%; overflow-x: auto; }}
.stTabs [role="tab"] {{ border-radius: 10px; padding: .55rem 1.05rem; color: var(--ink2); height: auto;
  transition: background .25s, color .25s; }}
.stTabs [role="tab"]:hover {{ background: rgba(255,255,255,.04); color: var(--ink); }}
.stTabs [role="tab"][aria-selected="true"] {{ background: linear-gradient(135deg, rgba(59,130,246,.38), rgba(34,211,238,.18)) !important;
  color: white !important; box-shadow: 0 8px 22px -10px rgba(59,130,246,.8); }}
.stTabs [role="tab"] p {{ font-weight: 600; }}
.stTabs .react-aria-SelectionIndicator, .stTabs [data-baseweb="tab-highlight"], .stTabs [data-baseweb="tab-border"] {{ display: none !important; }}
.stTabs [role="tablist"] + div, .stTabs > div > div:first-child {{ border-bottom: 0 !important; }}

/* ---------- Divers ---------- */
[data-testid="stDataFrame"], [data-testid="stDataEditor"] {{ border-radius: 14px; overflow: hidden; border: 1px solid var(--glass-brd); }}
.section-h {{ display:flex; align-items:center; gap:.6rem; margin: 1.6rem 0 .7rem; }}
.section-h b {{ font-family:'Space Grotesk'; font-size:1.2rem; color:var(--ink); }}
.section-h span {{ flex:1; height:1px; background: linear-gradient(90deg, var(--glass-brd), transparent); }}
.status-banner {{ display:flex; gap:1rem; align-items:center; padding:1rem 1.2rem; border-radius:16px; margin:.9rem 0;
  border:1px solid; animation: rise .6s both; }}
.status-banner .big {{ font-size:1.6rem; }}
.status-banner b {{ font-family:'Space Grotesk'; font-size:1.05rem; display:block; }}
.status-banner span {{ color: var(--ink2); font-size:.88rem; }}
.status-ok {{ background: rgba(12,163,12,.10); border-color: rgba(12,163,12,.45); }}
.status-ok b {{ color: #5FD35F; }}
.status-warn {{ background: rgba(250,178,25,.08); border-color: rgba(250,178,25,.45); }}
.status-warn b {{ color: {WARNING}; }}
.status-ko {{ background: rgba(208,59,59,.10); border-color: rgba(208,59,59,.5); }}
.status-ko b {{ color: #F07070; }}
.foot {{ text-align:center; color:var(--ink3); font-size:.78rem; margin-top:3rem; padding-top:1.2rem;
  border-top:1px solid var(--glass-brd); }}
.foot b {{ color: var(--ink2); }}
.stAlert {{ border-radius: 14px; }}
</style>
"""


def inject_css():
    st.markdown(GLOBAL_CSS, unsafe_allow_html=True)


# =============================================================================
# EN-TÊTE 3D (camion-remorque animé, parallaxe souris)
# =============================================================================

def _cuboid(cls: str, w: int, h: int, d: int, faces: dict | None = None) -> str:
    faces = faces or {}
    names = ["front", "back", "right", "left", "top", "bottom"]
    inner = "".join(f'<div class="f {n}">{faces.get(n, "")}</div>' for n in names)
    return f'<div class="cub {cls}" style="--w:{w}px;--h:{h}px;--d:{d}px">{inner}</div>'


def hero_3d(stats: dict | None = None, badge: str = "TIMAR TANGER · DÉDOUANEMENT IMPORT",
            title: str = "Ventilation automatique", highlight: str = "Groupage ALSTOM Fez",
            desc: str = "Factures, packing lists, CMR et manifestes lus automatiquement (OCR inclus), chaque ligne "
                        "rapprochée de votre référentiel NGP, et un classeur Excel de contrôle prêt pour la déclaration."):
    side_logo = ('<div class="brand"><div class="tl">TIMAR</div>'
                 '<div class="ts">INTERNATIONAL TRANSPORT · TANGER</div></div>')
    trailer = _cuboid("trailer", 300, 118, 96, {"front": side_logo, "back": side_logo,
                                                  "left": '<div class="doors"></div>'})
    cab = _cuboid("cab", 78, 98, 96, {"right": '<div class="wind"></div><div class="grill"></div>',
                                       "front": '<div class="win-s"></div>', "back": '<div class="win-s"></div>'})
    wheels = "".join(
        f'<div class="wheel" style="left:{x}px;transform:translateZ({z}px)"></div>'
        for x in (22, 62, 230, 340) for z in (49, -49 - 1))
    s = stats or {}
    chips = "".join(f'<div class="stat"><b>{html.escape(str(v))}</b><span>{html.escape(k)}</span></div>'
                    for k, v in s.items())
    page = f"""
<!doctype html><html><head><meta charset="utf-8">
<link href="{FONTS}" rel="stylesheet">
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
html,body{{background:transparent;overflow:hidden;font-family:Inter,system-ui,sans-serif;color:{INK}}}
.hero{{position:relative;height:370px;border-radius:26px;overflow:hidden;display:grid;grid-template-columns:1.05fr 1fr;
  background:linear-gradient(135deg,rgba(18,32,64,.95),rgba(9,16,32,.92) 55%,rgba(10,38,56,.9));
  border:1px solid rgba(120,160,255,.18);box-shadow:0 30px 80px -30px rgba(0,0,0,.9)}}
canvas{{position:absolute;inset:0;width:100%;height:100%;opacity:.55}}
.left{{position:relative;z-index:2;padding:38px 0 30px 42px;display:flex;flex-direction:column;justify-content:center}}
.badge{{display:inline-flex;align-items:center;gap:8px;font-size:11.5px;letter-spacing:.14em;font-weight:600;color:{INK_2};
  padding:7px 13px;border-radius:999px;border:1px solid rgba(120,160,255,.25);background:rgba(59,130,246,.10);width:max-content}}
.badge i{{width:7px;height:7px;border-radius:50%;background:#22D3EE;box-shadow:0 0 10px #22D3EE;animation:blink 1.6s infinite}}
@keyframes blink{{50%{{opacity:.3}}}}
h1{{font-family:'Space Grotesk',Inter,sans-serif;font-size:40px;line-height:1.06;margin:16px 0 12px;letter-spacing:-.02em}}
h1 .g{{background:linear-gradient(90deg,#60A5FA,#22D3EE 50%,{GOLD});-webkit-background-clip:text;background-clip:text;
  color:transparent;background-size:200% 100%;animation:sh 6s linear infinite}}
@keyframes sh{{to{{background-position:200% 0}}}}
p{{color:{INK_2};font-size:14.5px;line-height:1.55;max-width:520px}}
.stats{{display:flex;gap:10px;margin-top:20px;flex-wrap:wrap}}
.stat{{padding:9px 14px;border-radius:14px;background:rgba(255,255,255,.04);border:1px solid rgba(120,160,255,.16);
  backdrop-filter:blur(8px);animation:up .8s both}}
.stat:nth-child(2){{animation-delay:.1s}}.stat:nth-child(3){{animation-delay:.2s}}.stat:nth-child(4){{animation-delay:.3s}}
.stat b{{display:block;font-family:'Space Grotesk';font-size:17px;color:white}}
.stat span{{font-size:11px;color:{INK_3};letter-spacing:.06em;text-transform:uppercase}}
@keyframes up{{from{{opacity:0;transform:translateY(12px)}}}}
.stage{{position:relative;z-index:1;perspective:1100px;perspective-origin:50% 35%}}
.scene{{position:absolute;left:50%;top:56%;width:0;height:0;transform-style:preserve-3d;
  transform:rotateX(calc(-16deg + var(--my,0deg))) rotateY(calc(-32deg + var(--mx,0deg)));transition:transform .25s ease-out}}
.float{{transform-style:preserve-3d;animation:float 4s ease-in-out infinite}}
@keyframes float{{50%{{transform:translateY(-6px)}}}}
.truck{{position:absolute;transform-style:preserve-3d;transform:translate3d(-215px,-150px,0)}}
.cub{{position:absolute;transform-style:preserve-3d;width:var(--w);height:var(--h)}}
.f{{position:absolute;left:0;top:0;backface-visibility:visible}}
.front,.back{{width:var(--w);height:var(--h)}}
.front{{transform:translateZ(calc(var(--d)/2))}}
.back{{transform:rotateY(180deg) translateZ(calc(var(--d)/2))}}
.right,.left{{width:var(--d);height:var(--h);left:calc((var(--w) - var(--d))/2)}}
.right{{transform:rotateY(90deg) translateZ(calc(var(--w)/2))}}
.left{{transform:rotateY(-90deg) translateZ(calc(var(--w)/2))}}
.top,.bottom{{width:var(--w);height:var(--d);top:calc((var(--h) - var(--d))/2)}}
.top{{transform:rotateX(90deg) translateZ(calc(var(--h)/2))}}
.bottom{{transform:rotateX(-90deg) translateZ(calc(var(--h)/2))}}
.trailer .f{{background:linear-gradient(180deg,#F4F7FC,#C9D4E6);border:1px solid #8FA3C4}}
.trailer .top{{background:linear-gradient(90deg,#DCE5F2,#F7F9FD)}}
.trailer .right,.trailer .left{{background:linear-gradient(180deg,#D5DEEC,#AEBDD6)}}
.trailer .front::after,.trailer .back::after{{content:"";position:absolute;left:0;right:0;bottom:0;height:12px;
  background:linear-gradient(90deg,#1F3864,#2563EB,#1F3864)}}
.brand{{position:absolute;inset:0;display:flex;flex-direction:column;justify-content:center;align-items:center}}
.tl{{font-family:'Space Grotesk';font-weight:700;font-size:40px;letter-spacing:.12em;color:#1F3864}}
.ts{{font-size:8.5px;letter-spacing:.22em;color:#3B5A8F;margin-top:2px;font-weight:600}}
.doors{{position:absolute;inset:8px;border-left:1px solid #8FA3C4;background:linear-gradient(90deg,transparent 49.5%,#8FA3C4 50%,transparent 50.5%)}}
.cab{{transform:translate3d(306px,20px,0)}}
.cab .f{{background:linear-gradient(180deg,#2F6FE4,#1E40AF);border:1px solid #1B3A8A}}
.cab .top{{background:#3B82F6}}
.wind{{position:absolute;left:8px;right:8px;top:8px;height:38px;border-radius:5px;
  background:linear-gradient(135deg,#BAE6FD,#38BDF8 45%,#0C4A6E);box-shadow:inset 0 0 8px rgba(255,255,255,.6)}}
.grill{{position:absolute;left:14px;right:14px;bottom:12px;height:20px;border-radius:3px;
  background:repeating-linear-gradient(0deg,#0B1530 0 3px,#3B5A8F 3px 5px)}}
.win-s{{position:absolute;right:8px;top:8px;width:34px;height:30px;border-radius:4px;background:linear-gradient(135deg,#BAE6FD,#0C4A6E)}}
.wheel{{position:absolute;top:104px;width:34px;height:34px;border-radius:50%;
  background:radial-gradient(circle,#CBD5E1 0 22%,#0B1020 24% 60%,#1E293B 62%);box-shadow:0 0 0 3px #0B1020}}
.wheel::after{{content:"";position:absolute;inset:6px;border-radius:50%;border:2px dashed rgba(203,213,225,.55);animation:spin .9s linear infinite}}
@keyframes spin{{to{{transform:rotate(360deg)}}}}
.road{{position:absolute;width:760px;height:220px;left:-380px;top:-110px;transform:rotateX(90deg) translateZ(-140px);
  background:linear-gradient(90deg,transparent,rgba(30,41,70,.9) 15%,rgba(30,41,70,.9) 85%,transparent),#0000;
  -webkit-mask-image:radial-gradient(ellipse at center,#000 45%,transparent 72%)}}
.road::before{{content:"";position:absolute;left:0;right:0;top:50%;height:6px;margin-top:-3px;
  background:repeating-linear-gradient(90deg,{GOLD} 0 46px,transparent 46px 92px);animation:lane .7s linear infinite;opacity:.85}}
@keyframes lane{{to{{background-position:92px 0}}}}
.shadow{{position:absolute;width:460px;height:150px;left:-230px;top:-75px;transform:rotateX(90deg) translateZ(-139px);
  background:radial-gradient(ellipse,rgba(0,0,0,.6),transparent 65%)}}
.glow{{position:absolute;right:-60px;top:-60px;width:360px;height:360px;border-radius:50%;
  background:radial-gradient(circle,rgba(59,130,246,.35),transparent 65%);filter:blur(10px)}}
.boxes{{position:absolute;transform-style:preserve-3d}}
@media (max-width:820px){{.hero{{grid-template-columns:1fr}}.stage{{display:none}}h1{{font-size:30px}}.left{{padding:28px}}}}
</style></head><body>
<div class="hero">
  <canvas id="c"></canvas><div class="glow"></div>
  <div class="left">
    <div class="badge"><i></i>{html.escape(badge)}</div>
    <h1>{html.escape(title)}<br><span class="g">{html.escape(highlight)}</span></h1>
    <p>{html.escape(desc)}</p>
    <div class="stats">{chips}</div>
  </div>
  <div class="stage" id="stage"><div class="scene" id="scene"><div class="float">
    <div class="shadow"></div><div class="road"></div>
    <div class="truck">{trailer}{cab}{wheels}</div>
  </div></div></div>
</div>
<script>
const sc=document.getElementById('scene'),hero=document.querySelector('.hero');
hero.addEventListener('mousemove',e=>{{const r=hero.getBoundingClientRect();
  const x=(e.clientX-r.left)/r.width-.5,y=(e.clientY-r.top)/r.height-.5;
  sc.style.setProperty('--mx',(x*28)+'deg');sc.style.setProperty('--my',(-y*12)+'deg');}});
hero.addEventListener('mouseleave',()=>{{sc.style.setProperty('--mx','0deg');sc.style.setProperty('--my','0deg');}});
// réseau de points animé
const c=document.getElementById('c'),x=c.getContext('2d');let W,H,P=[];
function rs(){{W=c.width=c.offsetWidth*devicePixelRatio;H=c.height=c.offsetHeight*devicePixelRatio;
  P=Array.from({{length:46}},()=>({{x:Math.random()*W,y:Math.random()*H,vx:(Math.random()-.5)*.35,vy:(Math.random()-.5)*.35}}));}}
rs();addEventListener('resize',rs);
(function f(){{x.clearRect(0,0,W,H);const d=130*devicePixelRatio;
  for(const p of P){{p.x+=p.vx;p.y+=p.vy;if(p.x<0||p.x>W)p.vx*=-1;if(p.y<0||p.y>H)p.vy*=-1;
    x.fillStyle='rgba(125,170,255,.7)';x.beginPath();x.arc(p.x,p.y,1.6*devicePixelRatio,0,7);x.fill();}}
  for(let i=0;i<P.length;i++)for(let j=i+1;j<P.length;j++){{const a=P[i],b=P[j],k=Math.hypot(a.x-b.x,a.y-b.y);
    if(k<d){{x.strokeStyle='rgba(96,165,250,'+(.18*(1-k/d))+')';x.lineWidth=devicePixelRatio;x.beginPath();x.moveTo(a.x,a.y);x.lineTo(b.x,b.y);x.stroke();}}}}
  requestAnimationFrame(f);}})();
</script></body></html>"""
    components.html(page, height=378)


# =============================================================================
# ÉTAPES DU PROCESSUS
# =============================================================================

def steps_strip(active: int, steps: list | None = None):
    """active : 0 = rien, 1 = fichiers déposés, 2 = traitement fait."""
    steps = steps or [("📥", "Importer", "Dossier PDF + référentiel NGP"),
                      ("🧠", "Analyser", "OCR, extraction, matching NGP"),
                      ("📊", "Exporter", "Classeur Excel 3 feuilles")]
    cells = []
    for i, (ico, t, d) in enumerate(steps):
        state = "done" if i < active else ("cur" if i == active else "")
        cells.append(f'<div class="stp {state}"><div class="n">{ico}</div><div><b>{i + 1}. {t}</b>'
                     f'<span>{d}</span></div></div>')
    st.markdown(f"""
<style>
.steps{{display:grid;grid-template-columns:repeat(3,1fr);gap:.8rem;margin:1rem 0 1.1rem}}
.stp{{display:flex;gap:.8rem;align-items:center;padding:.85rem 1rem;border-radius:16px;background:var(--glass);
  border:1px solid var(--glass-brd);position:relative;overflow:hidden;transition:all .3s}}
.stp .n{{width:40px;height:40px;border-radius:12px;display:grid;place-items:center;font-size:1.15rem;
  background:rgba(255,255,255,.04);border:1px solid var(--glass-brd);flex:none}}
.stp b{{display:block;color:var(--ink);font-size:.92rem}} .stp span{{color:var(--ink3);font-size:.78rem}}
.stp.cur{{border-color:rgba(59,130,246,.55);box-shadow:0 0 0 1px rgba(59,130,246,.25),0 14px 34px -16px rgba(59,130,246,.6)}}
.stp.cur .n{{background:linear-gradient(135deg,#2563EB,#22D3EE);border:0;animation:bob 1.8s ease-in-out infinite}}
.stp.done .n{{background:rgba(12,163,12,.18);border-color:rgba(12,163,12,.5)}}
.stp.done::after{{content:"✓";position:absolute;right:14px;top:50%;transform:translateY(-50%);color:#5FD35F;font-weight:800}}
.stp::before{{content:"";position:absolute;left:0;bottom:0;height:2px;width:0;background:linear-gradient(90deg,#2563EB,#22D3EE)}}
.stp.done::before{{width:100%}} .stp.cur::before{{width:55%;animation:load 1.8s ease-in-out infinite alternate}}
@keyframes bob{{50%{{transform:translateY(-3px) rotate(-4deg)}}}}
@keyframes load{{from{{width:15%}}to{{width:85%}}}}
@media (max-width:760px){{.steps{{grid-template-columns:1fr}}}}
</style><div class="steps">{''.join(cells)}</div>""", unsafe_allow_html=True)


def card_title(icon: str, title: str, subtitle: str, tags: list[str] | None = None):
    tag_html = "".join(f"<em>{html.escape(t)}</em>" for t in (tags or []))
    st.markdown(f'<div class="card-title"><div class="ico">{icon}</div><div><b>{html.escape(title)}</b>'
                f'<span>{html.escape(subtitle)}</span></div></div>'
                + (f'<div class="doc-tags">{tag_html}</div>' if tag_html else ""), unsafe_allow_html=True)


def section(title: str, icon: str = ""):
    st.markdown(f'<div class="section-h"><b>{icon} {html.escape(title)}</b><span></span></div>', unsafe_allow_html=True)


def status_banner(kind: str, icon: str, title: str, detail: str):
    st.markdown(f'<div class="status-banner status-{kind}"><div class="big">{icon}</div><div><b>{html.escape(title)}</b>'
                f'<span>{html.escape(detail)}</span></div></div>', unsafe_allow_html=True)


def sidebar_brand():
    st.markdown('<div class="side-brand"><div class="logo">T</div><div><b>TIMAR Tanger</b>'
                '<span>Customs · Groupage</span></div></div>', unsafe_allow_html=True)


def sidebar_pill(color: str, text: str):
    st.markdown(f'<div class="side-pill"><i style="background:{color};box-shadow:0 0 8px {color}"></i>{text}</div>',
                unsafe_allow_html=True)


def footer():
    st.markdown('<div class="foot"><b>TIMAR Tanger</b> · Import (ventilation groupage) &amp; Export (BL / factures) ALSTOM · '
                'Traitement 100 % local — aucun document n\'est envoyé sur Internet</div>', unsafe_allow_html=True)


# =============================================================================
# CARTES KPI 3D (compteurs animés + inclinaison suivant la souris)
# =============================================================================

def kpi_cards(cards: list[dict]):
    """cards : {icon, label, value (int|str), sub, tone: blue|cyan|gold|good|warn|bad, decimals}"""
    tones = {"blue": ("#3B82F6", "rgba(59,130,246,.16)"), "cyan": ("#22D3EE", "rgba(34,211,238,.14)"),
             "gold": (GOLD, "rgba(245,184,61,.14)"), "good": (GOOD, "rgba(12,163,12,.16)"),
             "warn": (WARNING, "rgba(250,178,25,.14)"), "bad": (CRITICAL, "rgba(208,59,59,.16)")}
    items = []
    for i, c in enumerate(cards):
        col, soft = tones.get(c.get("tone", "blue"))
        val = c["value"]
        num = isinstance(val, (int, float)) and not isinstance(val, bool)
        data = f'data-to="{val}" data-dec="{c.get("decimals", 0)}"' if num else ""
        shown = "0" if num else html.escape(str(val))
        items.append(f"""<div class="k" style="--c:{col};--s:{soft};animation-delay:{i * 70}ms"><div class="in">
  <div class="top"><div class="ic">{c['icon']}</div><div class="lb">{html.escape(c['label'])}</div></div>
  <div class="v" {data}>{shown}</div><div class="sb">{html.escape(c.get('sub', ''))}</div></div></div>""")
    n = len(cards)
    page = f"""<!doctype html><html><head><meta charset="utf-8"><link href="{FONTS}" rel="stylesheet">
<style>
*{{box-sizing:border-box;margin:0;padding:0}} html,body{{background:transparent;font-family:Inter,sans-serif;overflow:hidden}}
.g{{display:grid;grid-template-columns:repeat({min(n, 6)},1fr);gap:14px;padding:8px 4px 12px;perspective:900px}}
.k{{position:relative;border-radius:20px;transform-style:preserve-3d;transition:transform .18s ease-out;animation:in .7s both;cursor:default}}
@keyframes in{{from{{opacity:0;transform:translateY(18px) rotateX(-18deg)}}}}
.in{{height:100%;border-radius:20px;padding:16px 16px 14px;
  background:linear-gradient(160deg,rgba(26,40,70,.92),rgba(12,20,38,.95));border:1px solid rgba(120,160,255,.16);
  box-shadow:0 18px 40px -20px rgba(0,0,0,.9),inset 0 1px 0 rgba(255,255,255,.05);position:relative;overflow:hidden}}
.in::before{{content:"";position:absolute;inset:-1px;border-radius:20px;padding:1px;
  background:linear-gradient(140deg,var(--c),transparent 45%);-webkit-mask:linear-gradient(#000 0 0) content-box,linear-gradient(#000 0 0);
  -webkit-mask-composite:xor;mask-composite:exclude;opacity:.8}}
.in::after{{content:"";position:absolute;width:140px;height:140px;right:-50px;top:-50px;border-radius:50%;
  background:radial-gradient(circle,var(--s),transparent 70%)}}
.shine{{position:absolute;inset:0;border-radius:20px;pointer-events:none;
  background:radial-gradient(260px circle at var(--px,50%) var(--py,0%),rgba(255,255,255,.10),transparent 45%)}}
.top{{display:flex;align-items:center;gap:10px;transform:translateZ(30px)}}
.ic{{width:34px;height:34px;border-radius:11px;display:grid;place-items:center;font-size:16px;background:var(--s);
  border:1px solid rgba(255,255,255,.08)}}
.lb{{color:{INK_2};font-size:12.5px;font-weight:600;letter-spacing:.02em}}
.v{{font-family:'Space Grotesk',Inter;font-size:30px;font-weight:700;color:#fff;margin:12px 0 2px;transform:translateZ(45px);
  text-shadow:0 6px 24px var(--s);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}
.sb{{color:{INK_3};font-size:11.5px;transform:translateZ(20px);min-height:15px}}
@media (max-width:900px){{.g{{grid-template-columns:repeat(2,1fr)}}}}
</style></head><body><div class="g">{''.join(items)}</div>
<script>
const fmt=(v,d)=>v.toLocaleString('fr-FR',{{minimumFractionDigits:d,maximumFractionDigits:d}});
document.querySelectorAll('.v[data-to]').forEach(el=>{{const to=+el.dataset.to,d=+el.dataset.dec,t0=performance.now(),D=1300;
  (function s(t){{const p=Math.min(1,(t-t0)/D),e=1-Math.pow(1-p,3);el.textContent=fmt(to*e,d);if(p<1)requestAnimationFrame(s);}})(t0);}});
document.querySelectorAll('.k').forEach(k=>{{const sh=document.createElement('div');sh.className='shine';k.firstElementChild.appendChild(sh);
  k.addEventListener('mousemove',e=>{{const r=k.getBoundingClientRect(),x=(e.clientX-r.left)/r.width,y=(e.clientY-r.top)/r.height;
    k.style.transform=`rotateY(${{(x-.5)*16}}deg) rotateX(${{(.5-y)*14}}deg) translateZ(10px)`;
    sh.style.setProperty('--px',x*100+'%');sh.style.setProperty('--py',y*100+'%');}});
  k.addEventListener('mouseleave',()=>{{k.style.transform='';}});}});
</script></body></html>"""
    rows = 1 if n <= 6 else 2
    components.html(page, height=152 * rows)


# =============================================================================
# GRAPHIQUES (Altair, livré avec Streamlit)
# =============================================================================

def _chart_config(chart):
    return (chart.configure(background="transparent", font="Inter")
            .configure_view(stroke=None)
            .configure_axis(labelColor=INK_2, titleColor=INK_2, gridColor="rgba(154,168,194,.12)",
                            domainColor="rgba(154,168,194,.25)", tickColor="rgba(154,168,194,.25)",
                            labelFontSize=11, titleFontSize=11, titleFontWeight=500)
            .configure_legend(labelColor=INK_2, titleColor=INK_2))


def chart_value_by_supplier(df: pd.DataFrame):
    """Barres horizontales : Valeur HT par fournisseur, une seule teinte.
    Une devise = un panneau avec sa propre échelle (on ne compare jamais des EUR à des USD)."""
    data = (df.groupby(["Fournisseur / Expéditeur", "Devise"], dropna=False)["Valeur HT"].sum().reset_index())
    data["Devise"] = data["Devise"].replace("", "?").fillna("?")
    data["Fournisseur"] = data["Fournisseur / Expéditeur"].fillna("?").str.slice(0, 34)
    data["Libellé"] = data["Valeur HT"].map(lambda v: f"{v:,.2f}".replace(",", " ").replace(".", ",")) + " " + data["Devise"]
    base = alt.Chart().encode(
        y=alt.Y("Fournisseur:N", sort="-x", title=None, axis=alt.Axis(labelLimit=240)),
        x=alt.X("Valeur HT:Q", title=None, axis=alt.Axis(format=",.0f", tickCount=4),
                scale=alt.Scale(nice=True, padding=0), ),
        tooltip=[alt.Tooltip("Fournisseur / Expéditeur:N", title="Fournisseur"), alt.Tooltip("Devise:N"),
                 alt.Tooltip("Valeur HT:Q", format=",.2f")])
    bars = base.mark_bar(color=SERIES_1, cornerRadiusTopRight=4, cornerRadiusBottomRight=4, height={"band": .62})
    labels = base.mark_text(align="left", dx=6, color=INK_2, fontSize=11).encode(text="Libellé:N")
    devises = list(data["Devise"].unique())
    for dev in devises:
        sub = data[data["Devise"] == dev]
        if len(devises) > 1:
            st.markdown(f'<div style="color:{INK};font-weight:600;font-size:.85rem;margin:.2rem 0 -.3rem">'
                        f'{html.escape(dev)}</div>', unsafe_allow_html=True)
        chart = alt.layer(bars, labels, data=sub).properties(height=max(60, 40 * len(sub)))
        st.altair_chart(_chart_config(chart), use_container_width=True, theme=None)


def chart_ngp_status(df: pd.DataFrame):
    """Statut du contrôle NGP : couleurs de statut + icône + libellé (jamais la couleur seule)."""
    order = [("✅ OK", "OK", GOOD), ("⚠️ À vérifier", "À vérifier", WARNING), ("❌ ÉCART", "ÉCART", CRITICAL)]
    counts = df["Contrôle NGP"].value_counts()
    total = max(int(counts.sum()), 1)
    data = pd.DataFrame([{"Statut": lab, "Lignes": int(counts.get(key, 0)),
                          "Part": int(counts.get(key, 0)) / total} for lab, key, _ in order])
    base = alt.Chart(data).encode(
        y=alt.Y("Statut:N", sort=[o[0] for o in order], title=None),
        x=alt.X("Lignes:Q", title="Nombre de lignes", axis=alt.Axis(tickMinStep=1, format="d")),
        tooltip=[alt.Tooltip("Statut:N"), alt.Tooltip("Lignes:Q"), alt.Tooltip("Part:Q", format=".0%")])
    bars = base.mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4, height={"band": .62}).encode(
        color=alt.Color("Statut:N", scale=alt.Scale(domain=[o[0] for o in order], range=[o[2] for o in order]),
                        legend=None))
    labels = base.mark_text(align="left", dx=6, color=INK_2, fontSize=11).encode(
        text=alt.Text("Lignes:Q", format="d"))
    st.altair_chart(_chart_config((bars + labels).properties(height=150)), use_container_width=True, theme=None)


