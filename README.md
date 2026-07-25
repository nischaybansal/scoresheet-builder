# Scoresheet Builder (Streamlit)

An internal web app for SHL Development Centres. A consultant uploads (or pastes) a
competency framework, ticks which assessment tools are in play, and downloads a
ready-to-score Excel workbook — with the **Detailed Integration** tab already wired so every
tool's scores flow into it automatically.

- **Online tools** (OPQ, MQ, SJT) carry the paste-stens engine (paste stens, map each behaviour
  to a dimension, the 1–5 score fills in and links through).
- **Offline tools** (Case Study, Role Play, BEI) are blank, formatted scaffolds for the
  consultant to write rubric anchors / interview questions into.
- **Detailed Integration** lists every behaviour, pulls each selected tool's score, holds the
  integrated behaviour rating and an auto competency rating, plus a summary table.

---

## What's in the folder

```
scoresheet-streamlit/
├── streamlit_app.py     the app
├── generator.py         the Excel engine (unchanged, no web framework in it)
├── requirements.txt
└── .streamlit/config.toml   brand colours
```

---

## 1. Run it on your own machine

Needs Python 3.9+.

```bash
cd scoresheet-streamlit
pip install -r requirements.txt
streamlit run streamlit_app.py
```

It opens in your browser at **http://localhost:8501**.

> Just want the Excel without the app? `python generator.py` writes a sample workbook and the
> blank template next to the script.

---

## 2. Give the whole team a link (recommended — no install, trusted domain)

Streamlit Community Cloud hosts it for free at a `*.streamlit.app` URL, so teammates just click
a link — nothing to run locally, no ports to open.

1. Put this folder in a **GitHub** repo (public, or private if you have Streamlit access to
   private repos).
2. Go to **https://share.streamlit.io**, sign in with GitHub, click **New app**.
3. Pick the repo, set the main file to **`streamlit_app.py`**, and **Deploy**.
4. Share the `https://<your-app>.streamlit.app` URL with the team.

To restrict who can open it, use Streamlit Cloud's **app settings → Sharing** to limit viewers
to specific emails / your organisation.

*(If your org runs its own hosting, `streamlit run streamlit_app.py --server.port 8501
--server.address 0.0.0.0` behind your standard proxy works too.)*

---

## 3. Use it

1. **Framework** — download the template, fill the two columns (Competency, Behaviour — one
   behaviour per row, repeat or blank the competency), upload it. Or paste the two columns.
2. **Tools** — tick the tools used in this centre.
3. **Generate** — name the file if you like, click **Build scoresheet**, then **Download**.

---

## Presenting it (3-minute demo)

1. Download the template, show a filled framework.
2. Tick OPQ + Role Play + BEI, build, open the file.
3. On **OPQ**, paste a couple of stens and map a behaviour — watch the score fill.
4. Flip to **Detailed Integration** — the same scores are already there.

---

## Changing things later

- Tools and the dimension lists live in `generator.py` (`TOOLS` registry, `*_DIMS` / `SJT_SCALES`).
- Multi-dimension mapping per behaviour (online tools) is a planned v2 — today it's one per behaviour.
- Assessor Comments / client Report tabs are out of v1 and can be added to the generator later.
