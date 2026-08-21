<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=DM+Mono:wght@400;500&display=swap');

/* Theme variables */
:root{
  --bg: #0b0b12;
  --panel: #0d0d1a;
  --panel-border: #252540;
  --panel-strong-border: #1e1e38;
  --accent: #7b8cf5;
  --accent-faint: rgba(123,140,245,0.15);
  --accent-faint-border: rgba(123,140,245,0.35);
  --muted: #9faabf;
  --text: #e8e8f8;
  --highlight: #3ecf8e;
  --stat-bg: #0d0d1a;
  --yellow: #FFFF00;
}

/* Apply base font */
html, body, [class*="css"] { font-family: 'Syne', sans-serif; color: var(--text); background-color: var(--bg); }

/* Header/card */
.moa-header {
    background: var(--panel);
    border: 1px solid var(--panel-border);
    border-radius: 20px;
    padding: 32px 28px 24px;
    margin-bottom: 24px;
    box-shadow: 0 20px 60px rgba(0,0,0,0.6);
}
.moa-badge {
    display: inline-block;
    background: var(--accent-faint);
    border: 1px solid var(--accent-faint-border);
    border-radius: 100px;
    padding: 4px 14px;
    font-family: 'DM Mono', monospace;
    font-size: 10px;
    color: var(--accent);
    letter-spacing: 0.1em;
    text-transform: uppercase;
    margin-bottom: 12px;
}
.moa-title { font-size: 26px; font-weight: 800; color: var(--text); margin-bottom: 4px; }
.moa-title em { color: var(--accent); font-style: normal; }
.moa-sub { font-family: 'DM Mono', monospace; font-size: 11px; color: var(--text); line-height: 1.6; }

/* Stat box */
.stat-box {
    background: var(--stat-bg);
    border: 1px solid var(--panel-strong-border);
    border-radius: 12px;
    padding: 14px;
    text-align: center;
}
.stat-num { font-size: 28px; font-weight: 800; color: var(--accent); display: block; }
.stat-lbl { font-family: 'DM Mono', monospace; font-size: 10px; color: var(--muted); text-transform: uppercase; letter-spacing: .08em; }

/* Excel-ready card */
.excel-ready {
  background: rgba(62,207,142,0.08);
  border: 1px solid rgba(62,207,142,0.4);
  border-radius: 16px;
  padding: 20px 22px;
  margin: 16px 0;
  color: var(--text);
}

/* Small helpers for table-like widgets */
.table-cell {
  background: transparent;
  color: var(--text);
}

/* Highlight roster cells if needed (keeps existing logic in app.py) */
.highlight-yellow { background: var(--yellow); }

/* Make sure links/buttons have a consistent accent */
a, button, .stButton>button {
  color: var(--text);
}
</style>
