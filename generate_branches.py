#!/usr/bin/env python3
"""
Generate 40-branches.mp3 and 40-branches.html from 40-Branches_Chart.xlsx.

Audio structure per row:
  4 s silence → 880 Hz bell → "Column: Value" for each non-empty cell
  → brief pause → Vedic Name repeated once

HTML slideshow stays in sync with the audio via JavaScript timeupdate.

Run:  python generate_branches.py
"""

import subprocess, sys, shutil, io, json, re
from pathlib import Path

# ── auto-install Python packages ──────────────────────────────────────────────
def _pip(*pkgs):
    subprocess.check_call([sys.executable, "-m", "pip", "install", *pkgs])

# audioop was removed in Python 3.13; pydub needs audioop-lts as a replacement
if sys.version_info >= (3, 13):
    try:
        import audioop          # noqa: F401
    except ImportError:
        print("Installing audioop-lts …"); _pip("audioop-lts")

for _pkg, _mod in [("openpyxl", "openpyxl"), ("gtts", "gtts"), ("pydub", "pydub")]:
    try:
        __import__(_mod)
    except ImportError:
        print(f"Installing {_pkg} …"); _pip(_pkg)

# ── check ffmpeg (needed by pydub for MP3 export) ─────────────────────────────
if not shutil.which("ffmpeg"):
    sys.exit(
        "\nERROR: ffmpeg not found on PATH — required for MP3 export.\n"
        "  Windows : winget install --id Gyan.FFmpeg\n"
        "  macOS   : brew install ffmpeg\n"
        "  Ubuntu  : sudo apt install ffmpeg\n"
        "After installing, close and reopen this terminal, then run again.\n"
    )

# ── imports ───────────────────────────────────────────────────────────────────
import openpyxl
from gtts import gTTS
from pydub import AudioSegment
from pydub.generators import Sine

# ── config ────────────────────────────────────────────────────────────────────
XLSX               = Path("40-Branches_Chart.xlsx")
MP3_OUT            = Path("40-branches.mp3")
HTML_OUT           = Path("40-branches.html")
PRONUNCIATION_GUIDE = Path("pronunciation_guide.txt")

# ── pronunciation guide ───────────────────────────────────────────────────────
def _load_pronunciations() -> dict[str, str]:
    if not PRONUNCIATION_GUIDE.exists():
        return {}
    subs: dict[str, str] = {}
    for line in PRONUNCIATION_GUIDE.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if "=" in line:
            word, phon = line.split("=", 1)
            subs[word.strip()] = phon.strip()
    print(f"Loaded {len(subs)} pronunciation entries from {PRONUNCIATION_GUIDE.name}")
    return subs

PRONUNCIATIONS: dict[str, str] = _load_pronunciations()
# Sort keys longest-first so multi-word phrases match before single words
_PRON_KEYS = sorted(PRONUNCIATIONS.keys(), key=len, reverse=True)


def preprocess(text: str) -> str:
    """Apply pronunciation substitutions then fix word-number hyphens."""
    for word in _PRON_KEYS:
        text = text.replace(word, PRONUNCIATIONS[word])
    # "Shrauta-20" → "Shrauta, 20"  (hyphen between letter and digit)
    text = re.sub(r"(\w)-(\d)", r"\1, \2", text)
    return text


SILENCE_PRE_MS = 4_000   # silence before each row
BELL_HZ        = 880
BELL_DUR_MS    = 2_000   # total bell clip length
BELL_FADE_MS   = 1_800   # long fade-out simulates resonance
GAP_CELLS_MS   = 350     # pause between spoken cells within a row
GAP_REPEAT_MS  = 600     # pause before Vedic Name repeat at end of row


# ── helpers ───────────────────────────────────────────────────────────────────
def tts(text: str) -> AudioSegment:
    text = preprocess(text)
    buf = io.BytesIO()
    gTTS(text=text, lang="en", slow=False).write_to_fp(buf)
    buf.seek(0)
    return AudioSegment.from_mp3(buf)


BELL: AudioSegment = (
    Sine(BELL_HZ)
    .to_audio_segment(duration=BELL_DUR_MS, volume=-16)   # -16 dBFS = gentle
    .fade_in(15)
    .fade_out(BELL_FADE_MS)
)


def sil(ms: int) -> AudioSegment:
    return AudioSegment.silent(duration=ms)


def esc(s: str) -> str:
    return (
        (s or "")
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


# ── load spreadsheet ──────────────────────────────────────────────────────────
if not XLSX.exists():
    sys.exit(f"ERROR: {XLSX} not found — place it next to this script.")

wb = openpyxl.load_workbook(XLSX)
ws = wb.active

raw_headers = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
headers = [str(h).strip() if h else f"Col{i+1}" for i, h in enumerate(raw_headers)]

entries: list[list[str]] = []
for row in ws.iter_rows(min_row=2, values_only=True):
    vals = [str(v).strip() if v is not None else "" for v in row]
    if any(vals):
        entries.append(vals)

print(f"Loaded {len(entries)} entries from {XLSX.name}")
print(f"Columns: {headers}\n")


# ── build audio track ─────────────────────────────────────────────────────────
track = sil(0)
slide_times_s: list[float] = []   # audio offset (seconds) when each slide should appear

for idx, row in enumerate(entries):
    vedic = row[0] if row else ""
    print(f"  [{idx+1:2d}/{len(entries)}] {vedic}", flush=True)

    # 4 s silence, then slide advances when bell starts
    track += sil(SILENCE_PRE_MS)
    slide_times_s.append(len(track) / 1000.0)
    track += BELL

    # Speak "Column Name: Value" for every non-empty cell
    first_cell = True
    for col_name, val in zip(headers, row):
        if not val:
            continue                   # skip column name too when cell is empty
        if not first_cell:
            track += sil(GAP_CELLS_MS)
        track += tts(f"{col_name}: {val}")
        first_cell = False

    # Repeat Vedic Name once after the full row
    if vedic:
        track += sil(GAP_REPEAT_MS)
        track += tts(vedic)

dur_s = len(track) / 1000.0
mins, secs = divmod(dur_s, 60)
print(f"\nTotal audio: {int(mins)}m {secs:.0f}s")
print(f"Exporting {MP3_OUT.name} …", flush=True)
track.export(str(MP3_OUT), format="mp3", bitrate="128k")
print(f"Saved: {MP3_OUT}")


# ── build HTML slideshow ──────────────────────────────────────────────────────
slides_html: list[str] = []
for i, row in enumerate(entries):
    vedic   = esc(row[0]) if len(row) > 0 else ""
    quality = esc(row[1]) if len(row) > 1 else ""
    aspect  = esc(row[2]) if len(row) > 2 else ""
    # Traditional Definition (col 3) is spoken but not shown on slide per spec

    inner = f'<p class="vedic">{vedic}</p>'
    if quality:
        inner += f'<p class="quality">{quality}</p>'
    if aspect:
        inner += f'<p class="aspect">{aspect}</p>'

    slides_html.append(f'  <div class="slide" id="s{i}">{inner}</div>')

SLIDES_BLOCK = "\n".join(slides_html)
TIMES_JSON   = json.dumps(slide_times_s)
N            = len(entries)
MP3_NAME     = MP3_OUT.name

html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0">
<title>40 Branches of the Veda</title>
<style>
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
html, body {{
  height: 100%;
  overflow: hidden;
  background: #0d1117;
  font-family: Georgia, 'Times New Roman', serif;
}}
#wrap {{
  display: flex;
  flex-direction: column;
  height: 100vh;
}}
#stage {{
  flex: 1;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 20px 24px 8px;
  overflow: hidden;
}}

/* ── slides ── */
.slide {{
  display: none;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  text-align: center;
  gap: 16px;
  max-width: 900px;
  width: 100%;
  animation: fi .5s ease;
}}
.slide.active {{ display: flex; }}
@keyframes fi {{
  from {{ opacity: 0; transform: translateY(8px); }}
  to   {{ opacity: 1; transform: translateY(0); }}
}}

/* ── text styles ── */
.vedic {{
  color: #f5e6c8;                                  /* warm cream */
  font-size: clamp(2rem, 6vw, 3.6rem);
  font-weight: bold;
  letter-spacing: .04em;
  line-height: 1.2;
}}
.quality {{
  color: #7eb8d4;                                  /* soft blue */
  font-size: clamp(1.15rem, 3.4vw, 2.1rem);
  font-style: italic;
  line-height: 1.35;
}}
.aspect {{
  color: #7dc47d;                                  /* soft green */
  font-size: clamp(.92rem, 2.5vw, 1.35rem);
  line-height: 1.5;
  max-width: 740px;
}}

/* ── idle prompt ── */
#idle {{
  color: #3a4760;
  font-size: clamp(1rem, 2.5vw, 1.25rem);
  letter-spacing: .03em;
}}

/* ── audio bar ── */
#bar {{
  display: flex;
  align-items: center;
  gap: 12px;
  background: rgba(10, 14, 20, .97);
  padding: 9px 16px;
  border-top: 1px solid #1e2a3a;
}}
#bar audio {{
  flex: 1;
  height: 36px;
  accent-color: #7eb8d4;
}}
#ctr {{
  color: #44506a;
  font-size: .78rem;
  font-family: monospace;
  white-space: nowrap;
  min-width: 56px;
  text-align: right;
}}
</style>
</head>
<body>
<div id="gate" style="position:fixed;top:0;left:0;width:100%;height:100%;background:#0d1117;display:flex;align-items:center;justify-content:center;z-index:1000;">
  <div style="text-align:center;color:#f5e6c8;font-family:Georgia,serif;">
    <p style="font-size:1.5rem;margin-bottom:1.2rem;">Enter password to continue</p>
    <input id="pwd" type="password" style="padding:8px 14px;font-size:1.1rem;border:1px solid #3a4760;background:#1a2232;color:#f5e6c8;border-radius:4px;outline:none;" placeholder="Password" autofocus />
    <br><br>
    <button onclick="checkPwd()" style="padding:8px 24px;font-size:1rem;background:#7eb8d4;color:#0d1117;border:none;border-radius:4px;cursor:pointer;">Enter</button>
    <p id="gate-err" style="display:none;color:#e07070;margin-top:1rem;">Incorrect password. Please try again.</p>
  </div>
</div>
<script>
function checkPwd() {{
  if (document.getElementById('pwd').value === 'vedic2026') {{
    document.getElementById('gate').style.display = 'none';
  }} else {{
    document.getElementById('gate-err').style.display = '';
    document.getElementById('pwd').value = '';
    document.getElementById('pwd').focus();
  }}
}}
document.getElementById('pwd').addEventListener('keydown', function(e) {{
  if (e.key === 'Enter') checkPwd();
}});
</script>
<div id="wrap">
  <div id="stage">
    <div id="idle">&#9654;&#xfe0e; Press play to begin</div>
{SLIDES_BLOCK}
  </div>
  <div id="bar">
    <audio id="aud" controls preload="metadata" src="{MP3_NAME}"></audio>
    <span id="ctr"></span>
  </div>
</div>

<script>
const TIMES = {TIMES_JSON};
const N     = {N};
const aud   = document.getElementById('aud');
const idle  = document.getElementById('idle');
const ctr   = document.getElementById('ctr');
let cur     = -1;

function show(idx) {{
  if (idx === cur) return;
  if (cur >= 0) document.getElementById('s' + cur).classList.remove('active');
  if (idx < 0) {{
    idle.style.display = '';
    ctr.textContent   = '';
  }} else {{
    idle.style.display = 'none';
    document.getElementById('s' + idx).classList.add('active');
    ctr.textContent = (idx + 1) + ' / ' + N;
  }}
  cur = idx;
}}

function sync() {{
  const t = aud.currentTime;
  let next = -1;
  for (let i = 0; i < TIMES.length; i++) if (t >= TIMES[i]) next = i;
  show(next);
}}

// After a seek, force-clear all active slides then re-sync cleanly
function forceSync() {{
  document.querySelectorAll('.slide.active')
          .forEach(el => el.classList.remove('active'));
  cur = -1;
  idle.style.display = '';
  sync();
}}

aud.addEventListener('timeupdate', sync);
aud.addEventListener('seeked',     forceSync);
aud.addEventListener('ended',      () => show(-1));

// ── playback speed ──────────────────────────────────────────────────────────
let speed = 1.0;
const MIN_SPEED = 0.25;
const MAX_SPEED = 3.0;

function setSpeed(s) {{
  speed = Math.min(MAX_SPEED, Math.max(MIN_SPEED, Math.round(s * 100) / 100));
  aud.playbackRate = speed;
  showBadge(speed.toFixed(2).replace(/\.?0+$/, '') + '×');
}}

// ── flash helpers ────────────────────────────────────────────────────────────
function showBadge(text) {{
  let b = document.getElementById('_badge');
  if (!b) {{
    b = document.createElement('div');
    b.id = '_badge';
    b.style.cssText = 'position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);'
      + 'background:rgba(0,0,0,.72);color:#f5e6c8;font-size:1.6rem;padding:12px 28px;'
      + 'border-radius:10px;pointer-events:none;z-index:600;font-family:Georgia,serif;'
      + 'transition:opacity .45s;';
    document.body.appendChild(b);
  }}
  b.textContent = text;
  b.style.opacity = '1';
  clearTimeout(b._t);
  b._t = setTimeout(() => b.style.opacity = '0', 1100);
}}

function showSeekFlash(side) {{
  const el = document.createElement('div');
  el.textContent = side === 'left' ? '◀ 10s' : '30s ▶';
  const pos = side === 'left' ? 'left:12%' : 'right:12%';
  el.style.cssText = 'position:fixed;top:50%;' + pos + ';transform:translateY(-50%);'
    + 'background:rgba(0,0,0,.65);color:#f5e6c8;font-size:1.3rem;padding:10px 22px;'
    + 'border-radius:36px;pointer-events:none;opacity:1;transition:opacity .5s;'
    + 'z-index:600;font-family:Georgia,serif;';
  document.body.appendChild(el);
  setTimeout(() => {{ el.style.opacity = '0'; setTimeout(() => el.remove(), 520); }}, 680);
}}

// ── keyboard shortcuts ───────────────────────────────────────────────────────
document.addEventListener('keydown', function(e) {{
  if (document.getElementById('gate').style.display !== 'none') return;
  switch (e.key) {{
    case ' ':
      e.preventDefault();
      aud.paused ? aud.play() : aud.pause();
      break;
    case 'ArrowRight':
      e.preventDefault();
      aud.currentTime = Math.min(aud.currentTime + 30, aud.duration || 0);
      showSeekFlash('right');
      break;
    case 'ArrowLeft':
      e.preventDefault();
      aud.currentTime = Math.max(aud.currentTime - 10, 0);
      showSeekFlash('left');
      break;
    case 'f': case 'F':
      e.preventDefault();
      setSpeed(speed + 0.25);
      break;
    case 's': case 'S':
      e.preventDefault();
      setSpeed(speed - 0.25);
      break;
  }}
}});

// ── tap zones (left = -10 s, right = +30 s) ──────────────────────────────────
document.getElementById('stage').addEventListener('click', function(e) {{
  if (document.getElementById('gate').style.display !== 'none') return;
  const rect = this.getBoundingClientRect();
  if (e.clientX < rect.left + rect.width / 2) {{
    aud.currentTime = Math.max(aud.currentTime - 10, 0);
    showSeekFlash('left');
  }} else {{
    aud.currentTime = Math.min(aud.currentTime + 30, aud.duration || 0);
    showSeekFlash('right');
  }}
}});
</script>
</body>
</html>"""

HTML_OUT.write_text(html_content, encoding="utf-8")
print(f"Saved: {HTML_OUT}")
print("\nAll done!")
print(f"  Audio  : {MP3_OUT.resolve()}")
print(f"  Slides : {HTML_OUT.resolve()}")
print("\nKeep both files in the same folder, then open the HTML in a browser.")
