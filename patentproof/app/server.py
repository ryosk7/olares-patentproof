"""特許明細書 AI校正 — Olares アプリ用の小さな Web サーバー（標準ライブラリのみ）

- /data 以下（Home/Documents をマウント）の .docx を一覧表示し、ボタンで校正を開始
- 校正は1件ずつ順番に実行（Qwen の同時実行数を超えないため）
- 結果は元ファイルと同じフォルダに【AI校正】〜.docx と 〜_修正箇所一覧.csv として保存
"""
from __future__ import annotations

import html
import os
import queue
import sys
import threading
import time
import traceback
import urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

sys.path.insert(0, os.path.dirname(__file__))
import proofread as pr  # noqa: E402

ROOT = Path(os.environ.get("DOCS_ROOT", "/data"))
PORT = int(os.environ.get("PORT", "8080"))
APP_DATA = Path(os.environ.get("APP_DATA", "/appdata"))
# 設定ファイルの置き場所（アプリ専用領域に書けない環境では Documents 内の隠しフォルダを使う）
CONFIG_CANDIDATES = [APP_DATA / "config.json", ROOT / ".patentproof" / "config.json"]

# 設定の既定値（インストール時の環境変数）。アプリの「設定」画面で上書きでき、上書きは APP_DATA/config.json に保存する
DEFAULTS = {
    "llm_base_url": os.environ.get("PROOF_LLM_BASE_URL") or pr.DEFAULT_BASE_URL,
    "llm_model": os.environ.get("PROOF_LLM_MODEL") or pr.DEFAULT_MODEL,
    "llm_api_key": os.environ.get("PROOF_LLM_API_KEY", ""),
    "reasoning": os.environ.get("PROOF_LLM_REASONING", "low"),
    "vote_runs": os.environ.get("PROOF_VOTE_RUNS", "3"),
    "parallel": os.environ.get("PROOF_LLM_PARALLEL", "2"),
    "ocr_url": os.environ.get("PROOF_OCR_URL", pr.OCR_URL),
    "manual": "",  # 空欄なら既定の校正方針（proofread.DEFAULT_MANUAL）
}


def load_config() -> dict:
    cfg = dict(DEFAULTS)
    try:
        import json
        for path in CONFIG_CANDIDATES:
            if path.exists():
                cfg.update({k: v for k, v in json.loads(path.read_text(encoding="utf-8")).items() if k in DEFAULTS})
                break
    except Exception:
        pass
    return cfg


def save_config(cfg: dict):
    import json
    data = json.dumps({k: cfg[k] for k in DEFAULTS}, ensure_ascii=False, indent=2)
    last = None
    for path in CONFIG_CANDIDATES:
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(data, encoding="utf-8")
            return
        except Exception as e:
            last = e
    raise last


def apply_config(cfg: dict):
    """設定を校正モジュールに反映する（ジョブ開始ごとに呼ぶ）"""
    pr.REASONING = cfg["reasoning"]
    pr.VOTE_RUNS = max(1, int(cfg["vote_runs"] or 3))
    pr.PARALLEL = max(1, int(cfg["parallel"] or 1))
    pr.OCR_URL = cfg["ocr_url"].strip()
    pr.MANUAL = (cfg.get("manual") or "").strip() or pr._load_manual()


def http_json(url: str, timeout=10, api_key: str = ""):
    import json
    import urllib.request
    req = urllib.request.Request(url, headers={"Authorization": "Bearer " + api_key} if api_key else {})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read() or b"{}")


def list_models(cfg: dict) -> list[str]:
    try:
        data = http_json(cfg["llm_base_url"].rstrip("/") + "/models", api_key=cfg["llm_api_key"])
        return sorted(m.get("id", "") for m in data.get("data", []) if m.get("id"))
    except Exception:
        return []


def check_status(cfg: dict) -> dict:
    st = {}
    models = list_models(cfg)
    if not models:
        st["llm"] = ("ng", "LLM に接続できません。設定の「LLM の URL」を確認してください。")
    elif cfg["llm_model"] not in models:
        st["llm"] = ("ng", f"モデル「{cfg['llm_model']}」が見つかりません（停止中か未インストール）。設定でモデルを選んでください。")
    else:
        st["llm"] = ("ok", "接続できます")
    if not cfg["ocr_url"].strip():
        st["ocr"] = ("muted", "未設定（図面のOCR照合は行いません）")
    else:
        try:
            import urllib.request
            urllib.request.urlopen(cfg["ocr_url"].rstrip("/") + "/health", timeout=5).read()
            st["ocr"] = ("ok", "接続できます")
        except Exception:
            st["ocr"] = ("ng", "OCR に接続できません（PaddleOCR が停止中、または URL 違い）。図面照合は省略されます。")
    return st

jobs: dict[str, dict] = {}
order: list[str] = []
q: "queue.Queue[str]" = queue.Queue()
lock = threading.Lock()


def safe(rel: str) -> Path:
    p = (ROOT / rel).resolve()
    if ROOT.resolve() not in p.parents and p != ROOT.resolve():
        raise ValueError("bad path")
    return p


def list_docx() -> list[Path]:
    out = []
    for dirpath, dirnames, filenames in os.walk(ROOT):
        dirnames[:] = [d for d in dirnames if not d.startswith(".")]
        if Path(dirpath).relative_to(ROOT).parts.__len__() > 3:
            continue
        for f in filenames:
            if f.lower().endswith(".docx") and not f.startswith("【AI校正") and not f.startswith("~$"):
                out.append(Path(dirpath) / f)
    return sorted(out, key=lambda p: p.stat().st_mtime, reverse=True)


def worker():
    while True:
        jid = q.get()
        job = jobs[jid]
        job["status"] = "実行中"
        job["started"] = time.time()

        def log(msg, job=job):
            job["log"].append(time.strftime("%H:%M:%S ") + msg)

        pr.log = log
        try:
            from docx import Document
            src = Path(job["path"])
            doc = Document(str(src))
            paras = pr.extract_paragraphs(doc)
            if not paras:
                raise RuntimeError("【書類名】明細書 以降の本文が見つかりません。")
            log(f"本文 {len(paras)} 段落 / {sum(len(p.text) for p in paras)} 文字")
            findings = pr.rule_checks(paras)
            log(f"ルールチェック: {len(findings)} 件")
            cfg = load_config()
            apply_config(cfg)
            llm = pr.LLM(cfg["llm_base_url"], cfg["llm_model"], cfg["llm_api_key"] or None, verbose=True)
            if job.get("mode") != "rules":
                findings += pr.llm_checks(llm, paras, int(os.environ.get("PROOF_SEGMENT_CHARS", "1200")))
            else:
                log("ルールのみモード（LLM は使いません）")
            pdfs = sorted(q for q in src.parent.glob("*.pdf") if "図" in q.name)
            if pdfs:
                log(f"図面PDF: {pdfs[0].name}")
                findings += pr.drawing_symbol_checks(paras, pdfs[0], allow_ocr=job.get("mode") == "rules")
            else:
                log("同じフォルダに図面PDF（名前に「図」を含む .pdf）がないため符号照合をスキップ")
            located = [f for f in findings if pr.locate(f, paras)]
            final = pr.dedupe(located)
            log(f"指摘 {len(final)} 件（原文に見つからない引用 {len(findings) - len(located)} 件を除外）")
            pr.apply_to_docx(doc, final)
            tag = "【AI校正・ルールのみ】" if job.get("mode") == "rules" else "【AI校正】"
            out_docx = src.parent / f"{tag}{src.name}"
            out_csv = src.parent / f"{src.stem}_修正箇所一覧{'_ルールのみ' if job.get('mode') == 'rules' else ''}.csv"
            doc.save(str(out_docx))
            import csv
            with open(out_csv, "w", newline="", encoding="utf-8-sig") as fp:
                w = csv.writer(fp)
                w.writerow(["No", "段落番号", "修正前", "修正後", "種類", "備考", "要確認", "反映", "検出元", "検出回数"])
                for i, f in enumerate(final, 1):
                    w.writerow([i, f.label, f.before, f.after, f.kind, f.note, "○" if f.needs_check else "", f.applied, f.source, f"{f.votes}/{pr.VOTE_RUNS}" if f.votes else "ルール"])
            job["outputs"] = [str(out_docx), str(out_csv)]
            job["rows"] = [(f.label, f.before, f.after, f.kind, f.note, f.applied) for f in final]
            job["status"] = "完了"
            log(f"LLM 呼び出し {llm.calls} 回")
        except SystemExit as e:
            job["status"] = "エラー"
            log(str(e))
        except Exception:
            job["status"] = "エラー"
            log(traceback.format_exc()[-1500:])
        job["finished"] = time.time()
        q.task_done()


CSS = """body{font-family:-apple-system,'Hiragino Sans','Noto Sans JP',sans-serif;margin:0;background:#f6f7f9;color:#1d1f23}
main{max-width:980px;margin:0 auto;padding:24px 16px}h1{font-size:20px;margin:0 0 4px}p.sub{color:#666;margin:0 0 20px;font-size:13px}
section{background:#fff;border:1px solid #e3e5e8;border-radius:10px;padding:16px;margin-bottom:16px}h2{font-size:15px;margin:0 0 12px}
table{width:100%;border-collapse:collapse;font-size:13px}td,th{border-bottom:1px solid #eee;padding:6px 8px;text-align:left;vertical-align:top}
button{background:#2f6fed;color:#fff;border:0;border-radius:6px;padding:6px 12px;cursor:pointer}button:disabled{background:#aab}
.st{font-weight:600}.ok{color:#0a7c3a}.ng{color:#c62828}.run{color:#b26a00}pre{white-space:pre-wrap;font-size:12px;background:#fafafa;padding:8px;border-radius:6px;max-height:220px;overflow:auto}
a{color:#2f6fed}.muted{color:#888;font-size:12px}"""


def rel(p) -> str:
    return str(Path(p).relative_to(ROOT))


def page_index() -> str:
    rows = []
    for p in list_docx()[:200]:
        r = rel(p)
        rows.append(f"<tr><td>{html.escape(r)}</td><td class=muted>{time.strftime('%Y-%m-%d %H:%M', time.localtime(p.stat().st_mtime))}</td>"
                    f"<td><form method=post action=/run style=display:inline><input type=hidden name=path value=\"{html.escape(r)}\"><button>校正する</button></form> <form method=post action=/run style=display:inline><input type=hidden name=path value=\"{html.escape(r)}\"><input type=hidden name=mode value=rules><button style=background:#667>ルールのみ</button></form></td></tr>")
    job_rows = []
    for jid in reversed(order[-30:]):
        j = jobs[jid]
        cls = {"完了": "ok", "エラー": "ng"}.get(j["status"], "run")
        el = ""
        if j.get("started"):
            el = f"{int((j.get('finished') or time.time()) - j['started'])}秒"
        job_rows.append(f"<tr><td><a href=/job/{jid}>{html.escape(rel(j['path']))}</a></td><td class='st {cls}'>{j['status']}</td><td class=muted>{el}</td></tr>")
    auto = "<meta http-equiv=refresh content=10>" if any(jobs[j]["status"] in ("待機中", "実行中") for j in order) else ""
    return f"""<!doctype html><meta charset=utf-8><meta name=viewport content="width=device-width,initial-scale=1">{auto}<title>明細書 AI校正</title><style>{CSS}</style>
<main><h1>明細書 AI校正</h1><p class=sub>Documents 内の明細書（.docx）を校正し、同じフォルダに「【AI校正】〜.docx」（変更履歴・コメント付き）と「〜_修正箇所一覧.csv」を保存します。1件あたり数分〜10分ほどかかります。</p>
<section><h2>実行状況</h2>{'<table><tr><th>ファイル</th><th>状態</th><th>経過</th></tr>' + ''.join(job_rows) + '</table>' if job_rows else '<p class=muted>まだ実行していません。</p>'}</section>
<section><h2>Documents 内の明細書</h2><table><tr><th>ファイル</th><th>更新日時</th><th></th></tr>{''.join(rows) or '<tr><td colspan=3 class=muted>.docx が見つかりません</td></tr>'}</table></section>
<p class=muted>LLM: {html.escape(load_config()["llm_model"])} ／ <a href=/settings>設定・接続確認</a></p></main>"""


def page_settings(msg: str = "") -> str:
    cfg = load_config()
    st = check_status(cfg)
    models = list_models(cfg)
    opts = "".join(f"<option value=\"{html.escape(m)}\" {'selected' if m == cfg['llm_model'] else ''}>{html.escape(m)}</option>" for m in models)
    if cfg["llm_model"] not in models:
        opts = f"<option value=\"{html.escape(cfg['llm_model'])}\" selected>{html.escape(cfg['llm_model'])}（見つかりません）</option>" + opts

    def row(label, inner, note=""):
        return f"<tr><th style=width:180px>{label}</th><td>{inner}<div class=muted>{note}</div></td></tr>"

    def inp(name, val, typ="text"):
        return f"<input name={name} type={typ} value=\"{html.escape(str(val))}\" style=width:100%>"

    sel = lambda name, val, choices: f"<select name={name}>" + "".join(f"<option {'selected' if c == val else ''}>{c}</option>" for c in choices) + "</select>"
    badge = lambda k: f"<span class='st {'ok' if st[k][0]=='ok' else 'ng' if st[k][0]=='ng' else 'muted'}'>{html.escape(st[k][1])}</span>"
    return f"""<!doctype html><meta charset=utf-8><meta name=viewport content="width=device-width,initial-scale=1"><title>設定</title><style>{CSS} input,select{{padding:6px;border:1px solid #ccd;border-radius:6px}}</style>
<main><p><a href=/>← 一覧へ</a></p><h1>設定・接続確認</h1><p class=sub>{html.escape(msg)}</p>
<section><h2>接続状態</h2><table>{row("LLM", badge("llm"))}{row("図面OCR", badge("ocr"))}</table></section>
<form method=post action=/settings><section><h2>LLM</h2><table>
{row("LLM の URL", inp("llm_base_url", cfg["llm_base_url"]), "OpenAI 互換 API の URL（例: Olares Router の共有エントランス http://xxxx.shared.olares.com/v1）")}
{row("モデル", f"<select name=llm_model style=width:100%>{opts}</select>", "上の URL から取得したモデル一覧。日本語に強い 27B 以上を推奨")}
{row("API キー", inp("llm_api_key", cfg["llm_api_key"], "password"), "Olares 内の LLM なら空欄でかまいません")}
{row("推論の深さ", sel("reasoning", cfg["reasoning"], ["low", "medium", "xhigh", ""]), "low が速く、xhigh は精度が上がる代わりに数倍遅い")}
{row("多数決の回数", sel("vote_runs", str(cfg["vote_runs"]), ["1", "3", "5"]), "3 推奨。全回一致の指摘だけ変更履歴にします")}
{row("同時実行数", sel("parallel", str(cfg["parallel"]), ["1", "2", "3", "4"]), "LLM 側の同時実行上限に合わせる")}
</table></section><section><h2>図面OCR（任意）</h2><table>
{row("OCR の URL", inp("ocr_url", cfg["ocr_url"]), "PaddleOCR の共有エントランス（設定 → アプリ → PaddleOCR → 共有エントランス）。空欄なら図面照合なし")}
</table><p class=muted>OCR は GPU を使うため、LLM と同時には動かしません。図面の読み取りは「ルールのみ」実行のときだけ行い、結果を図面PDFの横に保存してフル実行で再利用します。</p></section>
<section><h2>校正方針（任意）</h2>
<p class=muted>事務所の校正マニュアルなど、確認すること・しないことを箇条書きで書くと、すべてのチェックの指示に使われます。空欄なら既定の方針を使います。</p>
<textarea name=manual rows=12 style="width:100%;padding:6px;border:1px solid #ccd;border-radius:6px" placeholder="{html.escape(pr.DEFAULT_MANUAL)}">{html.escape(cfg.get("manual", ""))}</textarea></section>
<p><button>保存して接続確認</button></p></form></main>"""


def page_job(jid: str) -> str:
    j = jobs[jid]
    auto = "<meta http-equiv=refresh content=5>" if j["status"] in ("待機中", "実行中") else ""
    outs = "".join(f"<li>{html.escape(rel(o))}</li>" for o in j.get("outputs", []))
    rows = "".join(f"<tr><td>{html.escape(a)}</td><td>{html.escape(b)}</td><td>{html.escape(c)}</td><td>{html.escape(d)}</td><td>{html.escape(e)}</td><td class=muted>{html.escape(f)}</td></tr>"
                   for a, b, c, d, e, f in j.get("rows", []))
    table = f"<section><h2>修正箇所（{len(j.get('rows', []))}件）</h2><table><tr><th>段落</th><th>修正前</th><th>修正後</th><th>種類</th><th>備考</th><th>反映</th></tr>{rows}</table></section>" if rows else ""
    return f"""<!doctype html><meta charset=utf-8><meta name=viewport content="width=device-width,initial-scale=1">{auto}<title>校正結果</title><style>{CSS}</style>
<main><p><a href=/>← 一覧へ</a></p><h1>{html.escape(rel(j['path']))}</h1><p class=sub>状態: <b>{j['status']}</b></p>
{'<section><h2>保存したファイル（ファイルマネージャーで開けます）</h2><ul>' + outs + '</ul></section>' if outs else ''}
{table}<section><h2>ログ</h2><pre>{html.escape(chr(10).join(j['log'][-200:]))}</pre></section></main>"""


class H(BaseHTTPRequestHandler):
    def _send(self, body: str, code=200):
        b = body.encode()
        self.send_response(code)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(b)))
        self.end_headers()
        self.wfile.write(b)

    def do_GET(self):
        path = urllib.parse.urlparse(self.path).path
        if path == "/":
            return self._send(page_index())
        if path.startswith("/job/") and path[5:] in jobs:
            return self._send(page_job(path[5:]))
        if path == "/settings":
            return self._send(page_settings())
        if path == "/healthz":
            return self._send("ok")
        self._send("not found", 404)

    def do_POST(self):
        if self.path == "/settings":
            n = int(self.headers.get("Content-Length", "0"))
            form = urllib.parse.parse_qs(self.rfile.read(n).decode(), keep_blank_values=True)
            cfg = load_config()
            for k in DEFAULTS:
                if k in form:
                    cfg[k] = form[k][0].strip() if k != "manual" else form[k][0].replace("\r\n", "\n").strip()
            try:
                save_config(cfg)
                msg = "保存しました。"
            except Exception as e:
                msg = f"保存できませんでした: {e}"
            return self._send(page_settings(msg))
        if self.path != "/run":
            return self._send("not found", 404)
        n = int(self.headers.get("Content-Length", "0"))
        form = urllib.parse.parse_qs(self.rfile.read(n).decode())
        try:
            p = safe(form["path"][0])
        except Exception:
            return self._send("bad request", 400)
        jid = f"{int(time.time()*1000):x}"
        with lock:
            jobs[jid] = {"path": str(p), "status": "待機中", "log": [], "outputs": [], "mode": form.get("mode", ["full"])[0]}
            order.append(jid)
        q.put(jid)
        self.send_response(303)
        self.send_header("Location", f"/job/{jid}")
        self.end_headers()

    def log_message(self, *a):
        pass


if __name__ == "__main__":
    threading.Thread(target=worker, daemon=True).start()
    print(f"listening on :{PORT}, docs root {ROOT}", flush=True)
    ThreadingHTTPServer(("0.0.0.0", PORT), H).serve_forever()
