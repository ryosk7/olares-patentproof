# CLAUDE.md

特許明細書（.docx）をローカル LLM で事務校正する Olares アプリ。UI と校正対象は日本語。

## 構成
- `patentproof/` — Olares アプリのチャート（`OlaresManifest.yaml`, `Chart.yaml`, `values.yaml`, `templates/`）
- `patentproof/app/proofread.py` — 校正ロジック（CLI としても動く）
- `patentproof/app/server.py` — Web UI とジョブキュー（標準ライブラリの http.server、ワーカー1本で直列実行）
- `templates/deployment.yaml` — python:3.12-slim を使い、起動時に python-docx / pypdf を pip で入れる。アプリのコードは configmap で渡す
- `icon.png` — アプリアイコン（Manifest からは raw.githubusercontent.com の URL で参照している）

## 処理の流れ（proofread.py）
1. docx から本文の段落を取り出す（願書部分は対象外）
2. `rule_checks`: 符号の整合（R1a/R1b）、カタカナ表記ゆれ
3. `drawing_symbol_checks`: 図面PDF（ファイル名に「図」を含む）を PaddleOCR で読み、明細書→図面の向きで符号を照合する
   - OCR は rules モードのときだけ呼び、結果を `<pdf名>_OCR.txt` に保存する。full モードでは保存結果を使う
4. `llm_checks`: 観点別のプロンプトで LLM を `VOTE_RUNS` 回（既定3）サンプリングし、指摘を引用箇所でまとめて票数を数える
   - 全文で見る観点 `VIEWS_WHOLE`、区切りごとに見る観点 `VIEWS_LOCAL`
   - 2票以上は採用、1票のものは `sys_verify()` で検証する
5. `apply_to_docx`: 全回一致の小さな修正は変更履歴、それ以外は【要確認】コメントとして書き込み、CSV も出す

校正方針は `DEFAULT_MANUAL`。設定画面の「校正方針」欄、または `PROOF_MANUAL_FILE` で差し替えられ、`_sys_view` と `sys_verify()` の両方に入る。

## 守ること
- **事務所の校正マニュアルの文面や例をコードやリポジトリに入れない**。既定の方針は一般的な内容だけにする
- **評価用の正解データ（解答ファイル）の内容を、プロンプト、ルール、ナレッジに使わない**。正解は採点にだけ使う
- **特定の検証用文書の取りこぼしに合わせた個別ルールを足さない**。改善は汎用的なもの、または校正方針に基づくものに限る
- **LLM（Qwen）と PaddleOCR を同時に動かさない**。GPU／メモリが足りず LLM が落ちる。OCR は rules モードで先に済ませ、フル実行中は PaddleOCR を止める

## リリース手順
1. `Chart.yaml` の version と、`OlaresManifest.yaml` の `metadata.version` / `spec.versionName` を上げる
2. `tar czf patentproof-<ver>.tgz patentproof` でパッケージを作る（tgz は .gitignore 済み）
3. GitHub の Releases に tgz を添付する。Olares では Market → 自分のOlares → カスタムアプリパッケージをアップロード で更新する
- README（英語）とチャート内の `patentproof/README.md` は、画像行以外を同じ内容に保つ

## 動作確認
- `python3 -c "import ast; ast.parse(open('patentproof/app/proofread.py').read())"` で構文を確認する
- 本番の LLM／OCR は Olares の共有エントランス経由でしか届かない。ローカルではモックするか、Olares 上のアプリで確認する
