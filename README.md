# Patent Proofreader (patentproof) for Olares

<img src="icon.png" width="96" alt="">

An Olares app that proofreads Japanese patent specifications (.docx) with a local LLM running on your Olares.
Documents never leave your Olares.

> The proofreading itself targets **Japanese** patent documents, and the app UI is in Japanese.

## Features

- Pick a specification .docx from Home/Documents and proofread it
- The LLM checks each category (typos and mis-conversions / inconsistent terms / notation variants / colloquial or unnatural wording / reference-numeral mismatches) and runs 3 times; the number of votes becomes the confidence
- Small fixes found in all 3 runs become **tracked changes**; everything else becomes **comments** (marked 【要確認】, "please check"). Both are saved in the same folder:
  - `【AI校正】<original name>.docx`
  - `<original name>_修正箇所一覧.csv` (list of findings with vote counts)
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
2. To use drawing matching, first run **「ルールのみ」 (rules only)**, which takes a few seconds
   - The drawing is read with OCR and saved as `<drawing PDF name>_OCR.txt`
3. **Stop PaddleOCR**, then run a full check with **「校正する」 (Proofread)**. About 40 minutes for 20,000 characters
4. When finished, open `【AI校正】….docx` in the same folder with Word or OnlyOffice and review the tracked changes and comments

> **GPU note**: running the LLM and OCR at the same time can exhaust GPU/memory and stop the LLM.
> Run OCR (rules only) and the full check separately, and stop PaddleOCR during the full check.

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
| Proofreading policy | (built-in policy) | Paste your firm's proofreading manual or checklist; it is used in the instructions for every check |

## Environment

- Olares 1.12.6 or later (amd64 / arm64)
- Internet access is needed on first start to install python-docx and pypdf from PyPI

## License

MIT
