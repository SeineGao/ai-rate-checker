# AI Rate Checker

Local heuristic estimate of AI-writing traces in Word / plain text. **Not** Turnitin, GPTZero, or any official academic verdict.  
Built for self-check (e.g. PolyU coursework): upload `.docx`, parse in the browser, no paid API.

## Features

- Drag-and-drop or choose a **Word (.docx)** file (primary path)
- Fallback: paste plain text
- **AI rate 0–100%**, risk level, feature breakdown, highlighted suspicious sentences
- Near-duplicate sentence pairs **within the same document** (not an external corpus)
- Fully client-side; text does not leave your machine

## Quick start

```bash
cd /workspace/ai-rate-checker
python3 -m http.server 8877
```

Open:

```
http://127.0.0.1:8877/
```

You can also open `index.html` directly; a local HTTP server is more reliable.

## How to use

1. Save your paper / assignment as **.docx** (convert legacy `.doc` first).
2. Drop it on the dashed box, or click to select.
3. Click **Check AI rate**.
4. Review the score, feature bars, suspicious sentences, and in-doc duplicates.
5. Or click **Load sample text**, or expand “paste plain text”.

Bundled sample: `sample-ai-ish.docx` (Chinese boilerplate-heavy prose for pipeline testing). Heuristics still detect Chinese AI set phrases even though the UI is English.

## Technical notes

| Item | Detail |
|------|--------|
| Parsing | [mammoth.js](https://github.com/mwilliamson/mammoth.js) (vendored: `mammoth.browser.min.js`) |
| Scoring | Pure JS heuristics: ZH/EN boilerplate phrases, sentence-length burstiness, TTR, connector density, paragraph evenness, list density, etc. |
| Network | No cloud model calls; mammoth is local so the app works offline |

## Limitations

- **Not** grounds for academic misconduct decisions.
- Short texts, formulaic human writing, and heavily human-edited AI drafts can be mis-scored.
- “Plagiarism check” only compares sentences **inside the current document**.
- Weights target common LLM style in Chinese and English; formal templates may score high.

## Layout

```
ai-rate-checker/
  index.html
  styles.css
  app.js
  mammoth.browser.min.js
  sample-ai-ish.docx
  README.md
```

## Disclaimer

Self-check tool only. No warranty on accuracy or outcomes based on its results.
