# CLAUDE.md

特許明細書（.docx）をローカル LLM で事務校正する Olares アプリ。UI と校正対象は日本語。

## 構成
- `patentproof/` — Olares アプリのチャート（`OlaresManifest.yaml`, `Chart.yaml`, `values.yaml`, `templates/`）
- `patentproof/app/proofread.py` — 校正ロジック（CLI としても動く）
- `patentproof/app/server.py` — Web UI とジョブキュー（標準ライブラリの http.server、ワーカー1本で直列実行）
- `templates/deployment.yaml` — python:3.12-slim を使い、起動時に python-docx / pypdf / janome を pip で入れる。アプリのコードは configmap で渡す
- `icon.png` — アプリアイコン（Manifest からは raw.githubusercontent.com の URL で参照している）

## 処理の流れ
server.py の worker が工程（`STEPS`）を1つずつ実行し、工程ごとの状態・所要時間を画面に出す。CPU の工程を先に済ませ、GPU の工程（OCR → LLM）は重ねない。
1. 本文の抽出: docx から本文の段落を取り出す（願書部分は対象外）
2. 機械チェック `rule_checks`:
   - `symbol_table` / `symbol_checks`: 正規表現で名称と符号の対応表を作り、矛盾を探す（R1a 同じ符号に別名称、R1b 1回だけの符号、R1c 数字の打ち間違い、R1d 【符号の説明】と不一致、R1e 符号の説明との過不足）。対応表は `_符号対応表.csv` にも出す
   - `repeat_checks`: 形態素解析（janome）で同じ語の重複、正規表現で同じ文の繰り返し
   - `hyoki_checks`: カタカナ表記ゆれ
3. 図面の照合 `drawing_symbol_checks`: 図面PDF（ファイル名に「図」を含む）を PaddleOCR で読み、明細書→図面の向きで符号を照合する。OCR 結果は `<pdf名>_OCR.txt` に保存し、フル実行では保存結果があればそれを使う
4. LLM 観点チェック `llm_vote`: `VIEWS_LOCAL` を区切りごとに `VOTE_RUNS` 回（既定3）サンプリングし、引用箇所でまとめて票数を数える。同じプロンプトを続けて投げ、サーバーのプロンプトキャッシュを効かせる
   - 全文の観点 `VIEWS_WHOLE` は既定で使わない（`WHOLE_VIEWS`／設定画面）。2万字級で1回1.5〜2分かかり、問題2・3で18回とも0件だったため
5. LLM 検証 `llm_verify`: 1票のものを `sys_verify()` で検証する。2票以上は採用
6. 書き出し `apply_to_docx`: 全回一致の小さな修正は変更履歴、それ以外は【要確認】コメントとして書き込み、CSV も出す

GPU の取り合い対策: OCR・埋め込み・LLM はどれも1本のワーカーが工程の順に呼ぶので、処理が重ならない。OCR の処理が終われば次の工程に進む（PaddleOCR のアプリを止める必要はない）。`proofread.BEFORE_CALL` の `gpu_gate` では中止（/cancel/<id>）を反映する。
LLM の呼び出しごとの秒数・トークン数・finish_reason は `LLM.stats` に記録し、ログに `LLM.summary()` を出す（ボトルネックの確認用）。

校正方針は `DEFAULT_MANUAL`。「データセット」画面で使用オンのマニュアルがあればそれを使い、`_sys_view` と `sys_verify()` の両方に入る。旧版の設定画面の校正方針は初回にデータセットへ移す。`PROOF_MANUAL_FILE` でも差し替えられる。

## データセット（マニュアル）
- 保存先は APP_DATA/dataset：index.json（一覧）、`<id>.txt`（本文）、`<id>.chunks.json`（チャンクと埋め込み）、`<id>.rules.json`（ルール）、`<id>.orig.<拡張子>`（元ファイル）
- 取り込みは2段階。アップロード時は `ds_ingest` で CPU の処理だけ（文字の取り出し・`chunk_text`・書式からのルール読み取り）をしてすぐ返す。GPU の処理は `run_ds_task` を校正と同じワーカーのキューに入れ、1件ずつ行う
  - 順序：画像の OCR（PaddleOCR が動いているときだけ。PDF は文字のないページを `OCR_PAGES` ずつ）→ 埋め込み（`EMBED_BATCH` ずつ）→ LLM でのルール読み取り（`llm_manual_rules`）
- チャンクには項目番号（［M1-3］）を振る。校正では、使用中のマニュアルが `MANUAL_WARN_CHARS` 字以下なら全文、超えたら区切りごとに埋め込みで近い `top_k` 項目だけを `MANUAL_FOR` で渡す（設定で全文・検索に固定できる）
- 指摘の備考に根拠の項目番号を書かせ、結果画面ではツールチップ、docx のコメントと CSV では項目の冒頭を添える。根拠つきで変更履歴にした修正には【根拠】コメントも付ける
- ルール（`RULE_TYPES`）：replace / forbid / regex / ignore / symbol_exclude / symbol_prefix
  - 書式からは `parse_manual_rules` で読み取る（矢印、かぎ括弧の「とする」「ではなく」「使わない」「指摘しない」、誤／正の表）
  - LLM からは `llm_manual_rules` で読み取る。LLM が読み取ったルールの指摘は要確認にする
  - `validate_rule` で、入れ子の繰り返しや何にでも当たる正規表現をはじく
  - 校正では `manual_rule_checks` で当てはめ、symbol_exclude / symbol_prefix は `configure_symbols` で符号の抽出に足し、ignore は `is_ignored` で指摘を外す

## 守ること
- **事務所の校正マニュアルの文面や例をコードやリポジトリに入れない**。既定の方針は一般的な内容だけにする
- **評価用の正解データ（解答ファイル）の内容を、プロンプト、ルール、ナレッジに使わない**。正解は採点にだけ使う
- **特定の検証用文書の取りこぼしに合わせた個別ルールを足さない**。改善は汎用的なもの、または校正方針に基づくものに限る
- **LLM（Qwen）と PaddleOCR の処理を重ねない**。GPU／メモリが足りず LLM が落ちる。OCR は LLM の工程より前に済ませ、OCR が終わってから次の工程に進む（ワーカー1本で順番に呼ぶ）

## リリース手順
1. `Chart.yaml` の version と、`OlaresManifest.yaml` の `metadata.version` / `spec.versionName` を上げる
2. `tar czf patentproof-<ver>.tgz patentproof` でパッケージを作る（tgz は .gitignore 済み）
3. GitHub の Releases に tgz を添付する。Olares では Market → 自分のOlares → カスタムアプリパッケージをアップロード で更新する
- README（英語）とチャート内の `patentproof/README.md` は、画像行以外を同じ内容に保つ

## 動作確認
- `python3 -c "import ast; ast.parse(open('patentproof/app/proofread.py').read())"` で構文を確認する
- 本番の LLM／OCR は Olares の共有エントランス経由でしか届かない。ローカルではモックするか、Olares 上のアプリで確認する
