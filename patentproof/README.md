# Patent Proofreader (patentproof) for Olares

An Olares app that proofreads Japanese patent specifications (.docx) with a local LLM running on your Olares.
Documents never leave your Olares.

> The proofreading itself targets **Japanese** patent documents, and the app UI is in Japanese.

## Features

- Pick a specification .docx from Home/Documents and proofread it, or add one from the app by choosing files or drag & drop
- Runs in steps, one at a time: extract text → mechanical checks → drawing matching → LLM checks → LLM verification → write results. The job page shows each step's progress and time
- Mechanical checks (no LLM, a few seconds):
  - **Duplicated words** found with morphological analysis (janome), e.g. 「のの」「装置装置」「受け取って受け取った」, and sentences written twice
  - **Name ↔ reference-numeral table** built with regular expressions, flagging contradictions: one numeral with different names, a name whose numeral looks mistyped (dropped/doubled/swapped digit), names that differ from 【符号の説明】, numerals missing from either side, numerals used only once
  - Katakana notation variants (e.g. コンピュータ / コンピューター)
- The LLM checks each passage for typos/mis-conversions and colloquial or unnatural wording, 3 times each; the number of votes becomes the confidence
- Small fixes found in all 3 runs become **tracked changes**; everything else becomes **comments** (marked 【要確認】, "please check"). Both are saved in the same folder:
  - `【AI校正】<original name>.docx`
  - `<original name>_修正箇所一覧.csv` (list of findings with vote counts)
  - `<original name>_符号対応表.csv` (name ↔ reference-numeral table with contradictions)
- If a drawing PDF (file name containing 「図」) is in the same folder, reference numerals are read with OCR and matched against the specification (optional)

## Requirements

| Required? | App | Purpose |
|---|---|---|
| Required | An LLM reachable through Olares Router (e.g. "Qwen3.8-27B (llama.cpp)" from Market) | Proofreading |
| Optional | PaddleOCR (Market) | Matching reference numerals in drawing PDFs |

A model of 27B or larger with good Japanese ability is recommended.

## Installation

1. Download `patentproof-<version>.tgz` from [Releases](https://github.com/ryosk7/olares-patentproof/releases) (or build it from this repository with `tar czf patentproof-<version>.tgz patentproof`)
2. In Olares, open **Market → My Olares → Upload custom app package** and select the tgz
3. After installing, open the app and go to **「設定・接続確認」 (Settings & connection check)**
   - **LLM URL**: the URL shown under Settings → Applications → (your LLM app) → Shared entrance, with `/v1` appended
   - **Model**: after saving the URL, pick a model from the list
   - **OCR URL** (optional): Settings → Applications → PaddleOCR → Shared entrance
   - Click 「保存して接続確認」 (Save and check); you are ready when both show 「接続できます」 (connected)

Install-time defaults can also be changed through Olares environment variables (Settings → Applications → Patent Proofreader → Environment variables).
Values saved on the app's Settings page take precedence.

## Usage

1. Put the specification .docx (and drawing PDF) in any folder under Home/Documents
   - Or click 「ファイルを追加」 (Add files) in the app and choose them, or drag & drop them anywhere on the page. They are saved to `Home/Documents/アップロード/<specification name>/` (files with the same name are overwritten)
2. Click **「校正」 (Proofread)**. About 30 minutes for 20,000 characters (estimated from earlier run logs)
   - If the drawing has not been read yet, start PaddleOCR first. The drawing is read with OCR and saved as `<drawing PDF name>_OCR.txt`
   - After that, the job waits with 「PaddleOCR の停止待ち」 (waiting for PaddleOCR to stop). **Stop PaddleOCR** and the LLM steps start automatically
   - 「ルールのみ」 (rules only) runs just the mechanical checks and drawing matching, in a few seconds
3. A running job can be stopped with 「中止」 (Stop)
4. When finished, open `【AI校正】….docx` in the same folder with Word or OnlyOffice and review the tracked changes and comments
   - You can also download the result with 「結果」 (Result) in the list, or the .docx and CSV from the job's detail page

> **GPU note**: running the LLM and OCR at the same time can exhaust GPU/memory and stop the LLM.
> The app never calls the LLM while PaddleOCR is running: it checks before every LLM call and waits until PaddleOCR is stopped.

## Accuracy (for reference)

On a ~20,000-character test specification with 18 errors confirmed by hand, it found about 60% of them (Qwen3.8-27B, reasoning: low, 3-run voting).
Almost all findings applied as tracked changes were valid; comments need to be triaged by a person.
Always have a person do the final review.

## Settings

| Setting | Default | Description |
|---|---|---|
| Reasoning depth | low | medium / xhigh is more accurate but several times slower |
| Voting runs | 3 | Only findings that appear in every run become tracked changes |
| Parallel requests | 2 | Match the LLM's concurrency limit |
| Whole-document LLM views | off | Passes the whole specification to the LLM for term/notation/numeral consistency. Off because each call took 1.5–2 minutes and found nothing on 20,000-character specifications; the mechanical checks cover these instead |

### Dataset (proofreading manuals)

Manage your firm's proofreading manuals on the 「データセット」 (Dataset) page: add them by choosing files or drag & drop (.txt / .md / .docx / .pdf), edit, download, turn on/off, or delete.
Importing a file with the same name updates it. Manuals that are on are joined into the proofreading policy used in every LLM check; with none on, the built-in policy is used.
A policy written on the Settings page in older versions is moved here automatically. Long manuals slow down every LLM call, so keep only what matters for proofreading.

## Environment

- Olares 1.12.6 or later (amd64 / arm64)
- Internet access is needed on first start to install python-docx, pypdf and janome from PyPI

## License

MIT
