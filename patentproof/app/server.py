"""Proofy（特許明細書の AI 校正）— Olares アプリ用の小さな Web サーバー（標準ライブラリのみ）

- /data 以下（Home/Documents をマウント）の .docx を一覧表示し、ボタンで校正を開始
- 画面は左のサイドバー（明細書／実行履歴／設定）と本文。一覧の状態は /api/rows を数秒ごとに読んで差し替える
- 画面からファイルを選ぶかドラッグ＆ドロップすると Documents/アップロード/<明細書名>/ に保存する
- 校正は1件ずつ、工程（STEPS）を順番に実行する。CPU の工程（抽出・機械チェック）→ GPU の工程（OCR → LLM）の順で、
  OCR（PaddleOCR）と LLM の処理は重ならない（OCR が終わってから次の工程に進む）
- 結果は元ファイルと同じフォルダに【AI校正】〜.docx、〜_修正箇所一覧.csv、〜_符号対応表.csv として保存
- 校正マニュアルは「データセット」（APP_DATA/dataset）で管理し、有効なものをつなげて校正方針にする
"""
from __future__ import annotations

import gc
import html
import json
import re
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
    "whole_views": os.environ.get("PROOF_WHOLE_VIEWS", "0"),
    "manual_mode": "auto",  # auto: 8000字までは全文、超えたら検索 / full: 常に全文 / search: 常に検索
    "embed_model": os.environ.get("PROOF_EMBED_MODEL", ""),  # 空なら LLM の URL のモデル一覧から自動で選ぶ
    "top_k": "6",
    "rule_llm": "1",  # マニュアルからのルール読み取りに LLM も使うか
    "manual": "",  # 旧版の設定画面の校正方針。データセットへ移したあとは空
}


def load_config() -> dict:
    cfg = dict(DEFAULTS)
    try:
        for path in CONFIG_CANDIDATES:
            if path.exists():
                cfg.update({k: v for k, v in json.loads(path.read_text(encoding="utf-8")).items() if k in DEFAULTS})
                break
    except Exception:
        pass
    return cfg


def save_config(cfg: dict):
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
    pr.WHOLE_VIEWS = cfg.get("whole_views") == "1"
    pr.MANUAL = ds_manual() or (cfg.get("manual") or "").strip() or pr._load_manual()


# ---- データセット（校正マニュアル） ----
# APP_DATA/dataset/
#   index.json          一覧（id, code=M1 などの項目番号の頭, name, enabled, chars, chunks, embedded, images, images_unread …）
#   <id>.txt            取り出した本文（画像の OCR 結果を含む）
#   <id>.chunks.json    チャンク [{cid: "M1-3", text, vec}]（vec は埋め込み。未計算なら null）
#   <id>.orig.<拡張子>   取り込んだ元ファイル（画像の読み直し用）
DATASET_CANDIDATES = [APP_DATA / "dataset", ROOT / ".patentproof" / "dataset"]
DATASET_EXTS = (".txt", ".md", ".docx", ".pdf", ".png", ".jpg", ".jpeg")
MANUAL_WARN_CHARS = 8000  # これを超えると全文では渡さず、検索で関係する項目だけを渡す（「自動」のとき）
CHUNK_CHARS = 400
OCR_PAGES = 4           # スキャン PDF を OCR に送るときの1回あたりのページ数
EMBED_BATCH = 8         # 埋め込みを1回に送るチャンク数
OCR_IMAGE_MAX = 10 * 1024 * 1024  # これより大きい画像は OCR に送らない
ds_lock = threading.RLock()  # 旧設定の移行で ds_items → ds_put と入れ子になるため


def ds_dir() -> Path:
    for d in DATASET_CANDIDATES:
        try:
            d.mkdir(parents=True, exist_ok=True)
            probe = d / ".probe"
            probe.write_text("")
            probe.unlink()
            return d
        except Exception:
            continue
    raise RuntimeError("データセットの保存先に書き込めません")


def ds_items() -> list[dict]:
    d = ds_dir()
    idx = d / "index.json"
    if not idx.exists():
        items: list[dict] = []
        write_atomic(idx, "[]")
        # 旧版で設定画面に書いた校正方針をデータセットへ移す
        cfg = load_config()
        if (cfg.get("manual") or "").strip():
            ds_put("設定画面の校正方針.txt", cfg["manual"].strip().encode("utf-8"), "設定画面から移行")
            cfg["manual"] = ""
            save_config(cfg)
            items = json.loads(idx.read_text(encoding="utf-8"))
        return items
    items = json.loads(idx.read_text(encoding="utf-8"))
    n = max([int(x["code"][1:]) for x in items if x.get("code")] or [0])
    for x in items:  # 項目番号の頭（M1, M2 …）がない旧データに振る
        if not x.get("code"):
            n += 1
            x["code"] = f"M{n}"
    return items


def write_atomic(path: Path, text: str):
    """一時ファイルに書いてから置き換える（書きかけのファイルを別のスレッドが読まないように）"""
    tmp = path.with_name(f".{path.name}.{threading.get_ident()}.tmp")
    tmp.write_text(text, encoding="utf-8")
    os.replace(tmp, path)


def ds_save(items: list[dict]):
    write_atomic(ds_dir() / "index.json", json.dumps(items, ensure_ascii=False, indent=2))


def ds_text(item_id: str) -> str:
    return (ds_dir() / f"{item_id}.txt").read_text(encoding="utf-8")


def ds_chunks(item_id: str) -> list[dict]:
    f = ds_dir() / f"{item_id}.chunks.json"
    return json.loads(f.read_text(encoding="utf-8")) if f.exists() else []


def _orig(item: dict) -> Path:
    return ds_dir() / f"{item['id']}.orig{Path(item['name']).suffix.lower()}"


HEAD_RE = re.compile(r"^(#{1,6}\s|第[0-9０-９一二三四五六七八九十百]+[章節条項]|[0-9０-９]+[.．、)）]|[（(][0-9０-９]+[)）]|[■□◆◇●○▼▲【])")


def chunk_text(text: str) -> list[str]:
    """マニュアルを項目（見出し・箇条書き）の切れ目で CHUNK_CHARS 字前後にまとめる。見出しは次のチャンクの頭に付ける"""
    chunks, cur, heads, head = [], [], 0, ""
    for line in (l.strip() for l in text.splitlines()):
        if not line:
            continue
        is_head = bool(HEAD_RE.match(line)) and len(line) <= 60
        size = sum(len(x) for x in cur)
        # 見出しだけのチャンクは作らない（見出しが続くときは次の本文とまとめる）
        if cur and len(cur) > heads and (is_head or size + len(line) > CHUNK_CHARS):
            chunks.append("\n".join(cur))
            cur, heads = ([head], 1) if head and not is_head else ([], 0)  # 続きのチャンクにも見出しを付ける
        if is_head:
            head = line
            heads += 1
        cur.append(line)
    if cur:
        chunks.append("\n".join(cur))
    return [c for c in chunks if c.strip()]


def embed_model(cfg: dict, models: list[str] | None = None) -> str:
    """埋め込みモデル。設定が空なら LLM の URL のモデル一覧から小さい埋め込みモデル（gemma 系を優先）を選ぶ"""
    if cfg.get("embed_model"):
        return cfg["embed_model"]
    models = [m for m in (list_models(cfg) if models is None else models) if "embed" in m.lower()]
    return next((m for m in models if "gemma" in m.lower()), models[0] if models else "")


def embed(texts: list[str], cfg: dict, on_progress=None) -> list[list[float]]:
    import urllib.request
    model = embed_model(cfg)
    if not model:
        raise RuntimeError("埋め込みモデルが見つかりません")
    out: list[list[float]] = []
    headers = {"Content-Type": "application/json"}
    if cfg.get("llm_api_key"):
        headers["Authorization"] = "Bearer " + cfg["llm_api_key"]
    for i in range(0, len(texts), EMBED_BATCH):
        body = json.dumps({"model": model, "input": texts[i:i + EMBED_BATCH]}).encode()
        req = urllib.request.Request(cfg["llm_base_url"].rstrip("/") + "/embeddings", data=body, headers=headers)
        with urllib.request.urlopen(req, timeout=300) as r:
            data = json.loads(r.read())["data"]
        out += [d["embedding"] for d in sorted(data, key=lambda d: d.get("index", 0))]
        if on_progress:
            on_progress(len(out), len(texts))
    return out


def extract_text(name: str, data: bytes, ocr: bool) -> tuple[str, int, int]:
    """マニュアルのファイルから本文を取り出す。ocr=True なら画像（docx 内の画像・文字のない PDF ページ・画像ファイル）を OCR で読む。
    戻り値: (本文, 画像の数, 読めなかった画像の数)"""
    import io
    ext = Path(name).suffix.lower()
    if ext in (".txt", ".md"):
        for enc in ("utf-8-sig", "cp932", "euc-jp"):
            try:
                return data.decode(enc).replace("\r\n", "\n").strip(), 0, 0
            except UnicodeDecodeError:
                continue
        raise ValueError("文字コードを判別できません（UTF-8 か Shift_JIS で保存してください）")
    if ext in (".png", ".jpg", ".jpeg"):
        if not ocr:
            return "", 1, 1
        return "\n".join(pr.ocr_pages(data, 1)).strip(), 1, 0
    if ext == ".docx":
        from docx import Document
        from docx.oxml.ns import qn
        d = Document(io.BytesIO(data))
        lines, images, unread = [], 0, 0
        for p in d.paragraphs:
            lines.append(p.text)
            for blip in p._p.iter(qn("a:blip")):
                part = d.part.related_parts.get(blip.get(qn("r:embed")))
                if part is None:
                    continue
                images += 1
                if not ocr or len(part.blob) > OCR_IMAGE_MAX:
                    unread += 1
                    continue
                try:
                    got = "\n".join(pr.ocr_pages(part.blob, 1)).strip()
                    if got:
                        lines.append(f"［図の文字］\n{got}")
                except Exception:
                    unread += 1
        for t in d.tables:
            for row in t.rows:
                lines.append("｜".join(c.text.strip() for c in row.cells))
        return "\n".join(lines).strip(), images, unread
    if ext == ".pdf":
        from pypdf import PdfReader
        try:
            pages = [(pg.extract_text() or "").strip() for pg in PdfReader(io.BytesIO(data)).pages]
        except Exception as e:
            raise ValueError(f"PDF を読めませんでした（壊れているか、パスワード付きの可能性: {e}）")
        blank = [i for i, t in enumerate(pages) if len(t) < 20]  # 文字のないページ（スキャン・画像）
        if blank and ocr:
            # 文字のないページだけを OCR_PAGES ページずつ別の PDF にして送る（GPU のメモリを一度に使いすぎない）
            from pypdf import PdfWriter
            reader = PdfReader(io.BytesIO(data))
            for k in range(0, len(blank), OCR_PAGES):
                part = blank[k:k + OCR_PAGES]
                w = PdfWriter()
                for i in part:
                    w.add_page(reader.pages[i])
                buf = io.BytesIO()
                w.write(buf)
                for i, t in zip(part, pr.ocr_pages(buf.getvalue(), 0)):
                    pages[i] = t.strip()
                del w, buf
                gc.collect()
            return "\n".join(pages).strip(), len(blank), 0
        return "\n".join(pages).strip(), len(blank), len(blank)
    raise ValueError("対応していない形式です（.txt / .md / .docx / .pdf / .png / .jpg）")


def ds_rules(item_id: str) -> list[dict]:
    f = ds_dir() / f"{item_id}.rules.json"
    return json.loads(f.read_text(encoding="utf-8")) if f.exists() else []


def _write_chunks_rules(item: dict, text: str):
    """本文 → チャンク分割 → 書式からのルール読み取り（すべて CPU）。LLM で読み取ったルールは次の LLM の処理で読み直す。
    版番号（rev）を上げ、それより前に始まったバックグラウンドの処理が古い結果を書き戻さないようにする"""
    d = ds_dir()
    write_atomic(d / f"{item['id']}.txt", text)
    chunks = [{"cid": f"{item['code']}-{i}", "text": c, "vec": None} for i, c in enumerate(chunk_text(text), 1)]
    old = {pr.rule_key(r): r.get("enabled", True) for r in ds_rules(item["id"])}
    rules = pr.parse_manual_rules(chunks)
    for r in rules:  # 画面で切ったルールは、読み直しても切ったままにする
        r["enabled"] = old.get(pr.rule_key(r), True)
    write_atomic(d / f"{item['id']}.chunks.json", json.dumps(chunks, ensure_ascii=False))
    write_atomic(d / f"{item['id']}.rules.json", json.dumps(rules, ensure_ascii=False, indent=1))
    item.update(chars=len(text), chunks=len(chunks), rules=len(rules), rules_llm=0, embedded=False, updated=time.time(),
                rev=item.get("rev", 0) + 1, note="", task="処理待ち", llm_done=False)
    return chunks


def ds_enqueue(item_id: str):
    q.put(("ds", item_id))


def ds_put(name: str, data: bytes, source: str) -> tuple[dict, bool]:
    """同じ名前があれば内容を差し替える（更新）。なければ追加する。戻り値: (項目, 更新したか)。
    文字の取り出しはロックの外で先に行い、取り出せなければ既存のものには触らない"""
    text, images, unread = extract_text(name, data, ocr=False)
    if not text and not images:
        raise ValueError("文字を取り出せませんでした")
    with ds_lock:
        items = ds_items()
        cur = next((x for x in items if x["name"] == name), None)
        if cur:
            item = cur
        else:
            n = max([int(x["code"][1:]) for x in items] or [0]) + 1
            item = {"id": f"{int(time.time() * 1000):x}", "code": f"M{n}", "name": name, "enabled": True}
            items.append(item)
        item.update(source=source, images=images, images_unread=unread)
        _orig(item).write_bytes(data)
        _write_chunks_rules(item, text)
        ds_save(items)
    ds_enqueue(item["id"])
    return item, bool(cur)


def ds_update(item_id: str, **kw):
    # 再処理で元ファイルから読み直すときは、解析をロックの外で先に行う
    pre = None
    if kw.get("reprocess"):
        item = next(x for x in ds_items() if x["id"] == item_id)
        if item.get("source") != "画面で編集" and _orig(item).exists():
            pre = extract_text(item["name"], _orig(item).read_bytes(), ocr=False)
    with ds_lock:
        items = ds_items()
        item = next(x for x in items if x["id"] == item_id)
        queue_it = False
        if "text" in kw:
            item.update(source="画面で編集", images_unread=0)  # 編集した本文を正とする（画像を読み直して上書きしない）
            _write_chunks_rules(item, kw["text"])
            queue_it = True
        if kw.get("reprocess"):
            if pre:
                item.update(images=pre[1], images_unread=pre[2])
                _write_chunks_rules(item, pre[0])
            else:
                _write_chunks_rules(item, ds_text(item_id))  # 編集した本文はそのまま、チャンク・ルール・埋め込みをやり直す
            queue_it = True
        if "enabled" in kw:
            item["enabled"] = kw["enabled"]
        if "rule_key" in kw:  # ルール1件の使用オン／オフ（並び順ではなく中身で指定する）
            rules = ds_rules(item_id)
            for r in rules:
                if pr.rule_key(r) == kw["rule_key"]:
                    r["enabled"] = kw["rule_on"]
            write_atomic(ds_dir() / f"{item_id}.rules.json", json.dumps(rules, ensure_ascii=False, indent=1))
        ds_save(items)
    if queue_it:
        ds_enqueue(item_id)


def ds_delete(item_id: str):
    with ds_lock:
        items = ds_items()
        item = next((x for x in items if x["id"] == item_id), None)
        if item:
            for f in (ds_dir() / f"{item_id}.txt", ds_dir() / f"{item_id}.chunks.json", ds_dir() / f"{item_id}.rules.json", _orig(item)):
                f.unlink(missing_ok=True)
        ds_save([x for x in items if x["id"] != item_id])


def ds_active_chunks() -> list[dict]:
    """使用オンのマニュアルのチャンク（name を付けて返す）"""
    out = []
    for x in ds_items():
        if x.get("enabled"):
            chunks = ds_chunks(x["id"])
            if not chunks and (ds_dir() / f"{x['id']}.txt").exists():  # チャンク分割より前の版で取り込んだもの
                chunks = [{"cid": f"{x['code']}-{i}", "text": c, "vec": None} for i, c in enumerate(chunk_text(ds_text(x["id"])), 1)]
            out += [dict(c, name=x["name"]) for c in chunks]
    return out


def ds_active_rules() -> list[dict]:
    """使用オンのマニュアルの、使用オンのルール"""
    out = []
    for x in ds_items():
        if x.get("enabled"):
            out += [r for r in ds_rules(x["id"]) if r.get("enabled", True)]
    return out


def numbered(chunks: list[dict]) -> str:
    """チャンクに項目番号を付けて校正方針の文にする（LLM が根拠として番号を書けるように）"""
    return "\n".join(f"［{c['cid']}］ {c['text']}" for c in chunks)


def ds_manual() -> str:
    """使用オンのマニュアル全文（項目番号つき）。なければ空"""
    try:
        return numbered(ds_active_chunks())
    except Exception:
        return ""


def _cos(a: list[float], b: list[float]) -> float:
    import math
    na, nb = math.sqrt(sum(x * x for x in a)), math.sqrt(sum(x * x for x in b))
    return sum(x * y for x, y in zip(a, b)) / (na * nb) if na and nb else 0.0


def http_json(url: str, timeout=10, api_key: str = ""):
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


UPLOAD_DIR = "アップロード"
UPLOAD_EXTS = (".docx", ".pdf")
UPLOAD_MAX = 200 * 1024 * 1024


def upload_name(name: str) -> str:
    """アップロードされたファイル名・フォルダ名からパス区切りなどを取り除く"""
    name = Path(name.replace("\\", "/")).name.strip()
    if not name or name.startswith(".") or name.startswith("~$"):
        raise ValueError("bad name")
    return name


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


STEPS = [("extract", "本文の抽出", "CPU"), ("rules", "機械チェック", "CPU"), ("drawing", "図面の照合", "GPU"),
         ("retrieve", "マニュアルの検索", "GPU"), ("vote", "LLM 観点チェック", "GPU"), ("verify", "LLM 検証", "GPU"),
         ("write", "書き出し", "CPU")]


class Cancelled(Exception):
    pass


_ocr_state = {"t": 0.0, "up": False}


def ocr_running(force: bool = False) -> bool:
    """PaddleOCR が動いている（GPU を使っている）か。30秒キャッシュする"""
    if not pr.OCR_URL:
        return False
    if force or time.time() - _ocr_state["t"] > 30:
        import urllib.request
        try:
            urllib.request.urlopen(pr.OCR_URL.rstrip("/") + "/health", timeout=3).read()
            _ocr_state["up"] = True
        except Exception:
            _ocr_state["up"] = False
        _ocr_state["t"] = time.time()
    return _ocr_state["up"]


def gpu_gate(job: dict, force: bool = False):
    """LLM を呼ぶ前に毎回呼ぶ。中止を反映する。
    GPU の取り合いは工程の順番で防ぐ: OCR（図面・マニュアルの画像）と LLM・埋め込みはどれもこのワーカーが1つずつ呼ぶので、
    OCR の処理が終わってから次の工程に進む（PaddleOCR のアプリを止める必要はない）"""
    if job.get("cancel"):
        raise Cancelled()


class step:
    """with step(job, "rules"): … で工程の状態と所要時間を記録する"""

    def __init__(self, job: dict, key: str):
        self.job, self.key = job, key

    def __enter__(self):
        if self.job.get("cancel"):
            raise Cancelled()
        self.st = self.job["steps"][self.key]
        self.st.update(status="実行中", started=time.time(), pct=0)
        self.job["current"] = self.key
        pr.PROGRESS = self.set  # 工程の中の処理（LLM の呼び出し・OCR・段落ごとの処理など）が進み具合を知らせる
        return self.st

    def set(self, done: int, total: int):
        self.st["pct"] = min(99, int(done * 100 / total)) if total else 0

    def __exit__(self, et, ev, tb):
        pr.PROGRESS = None
        self.st["secs"] = time.time() - self.st["started"]
        self.st["status"] = "完了" if et is None else "キャンセル" if et is Cancelled else "エラー"
        if et is None:
            self.st["pct"] = 100
        gc.collect()  # 工程ごとにメモリを返し、次の工程（LLM の並列呼び出しなど）と取り合わないようにする
        return False


class Stale(Exception):
    """処理中にマニュアルが更新・削除された（新しい処理がキューに入っているので、この結果は捨てる）"""


def is_embedded(x: dict, cfg: dict) -> bool:
    """今の埋め込みモデルで計算済みか（モデルを変えたら計算し直す）"""
    return bool(x.get("embedded")) and x.get("embed_model") == embed_model(cfg)


def ds_save_if_current(item_id: str, rev: int, files: dict[str, str] | None = None, **kw) -> dict:
    """マニュアルが rev のままなら、ファイルと一覧の項目を書く。変わっていれば Stale"""
    with ds_lock:
        items = ds_items()
        it = next((x for x in items if x["id"] == item_id), None)
        if it is None or it.get("rev", 0) != rev:
            raise Stale()
        for name, text in (files or {}).items():
            write_atomic(ds_dir() / name, text)
        it.update(kw)
        ds_save(items)
        return it


def ds_fill_vectors(cfg: dict) -> int:
    """使用オンのマニュアルで、今のモデルの埋め込みがないものを計算して保存する（計算はロックの外）。計算した数を返す"""
    n = 0
    model = embed_model(cfg)
    for x in ds_items():
        if not x.get("enabled") or is_embedded(x, cfg):
            continue
        chunks = ds_chunks(x["id"])
        if not chunks and (ds_dir() / f"{x['id']}.txt").exists():  # チャンク分割より前の版で取り込んだもの
            with ds_lock:
                items = ds_items()
                it = next(y for y in items if y["id"] == x["id"])
                chunks = _write_chunks_rules(it, ds_text(x["id"]))
                ds_save(items)
                x = it
        if not chunks:
            continue
        for c, v in zip(chunks, embed([c["text"] for c in chunks], cfg)):
            c["vec"] = v
        try:
            ds_save_if_current(x["id"], x.get("rev", 0), {f"{x['id']}.chunks.json": json.dumps(chunks, ensure_ascii=False)},
                               embedded=True, embed_model=model)
            n += len(chunks)
        except Stale:
            pass
    return n


def run_ds_task(item_id: str):
    """マニュアルの GPU の処理（画像の OCR → 埋め込み → LLM でのルール読み取り）。
    校正と同じワーカーで1件ずつ、段階ごとに順番に動くので、OCR・埋め込み・LLM が同時に GPU を使わない。
    段階ごとに保存し、メモリはその都度返す。途中でマニュアルが更新・削除されたら、古い結果は書かずにやめる"""
    def update(**kw):
        return ds_save_if_current(item_id, rev, **kw)

    notes = []
    pr.log = lambda m: None
    try:
        item = next((x for x in ds_items() if x["id"] == item_id), None)
        if item is None:
            return
        rev = item.get("rev", 0)
        cfg = load_config()
        apply_config(cfg)
        # 1) 画像の OCR（PaddleOCR が動いているときだけ。画面で編集したものは編集した本文を正とするので読み直さない）
        if item.get("images_unread") and item.get("source") != "画面で編集" and _orig(item).exists():
            if ocr_running(force=True):
                update(task=f"画像の読み取り中（{item['images_unread']} 件）")
                text, images, unread = extract_text(item["name"], _orig(item).read_bytes(), ocr=True)
                with ds_lock:
                    items = ds_items()
                    it = next((x for x in items if x["id"] == item_id), None)
                    if it is None or it.get("rev", 0) != rev:
                        raise Stale()
                    it.update(images=images, images_unread=unread)
                    _write_chunks_rules(it, text)
                    it["task"] = "処理中"
                    rev = it["rev"]
                    ds_save(items)
                gc.collect()
            else:
                notes.append("PaddleOCR が停止中のため画像は未読み取り（起動してから「再処理」）")

        # 2) 埋め込み（EMBED_BATCH 件ずつ。計算はロックの外）
        chunks = ds_chunks(item_id)
        if chunks:
            update(task="埋め込みを計算中")
            try:
                for c, v in zip(chunks, embed([c["text"] for c in chunks], cfg)):
                    c["vec"] = v
                update(files={f"{item_id}.chunks.json": json.dumps(chunks, ensure_ascii=False)}, embedded=True, embed_model=embed_model(cfg))
            except Stale:
                raise
            except Exception as e:
                notes.append(f"埋め込みは未計算（{str(e)[:80]}）")
            del chunks
            gc.collect()

        # 3) LLM でのルール読み取り（書式から読み取れない自由な書き方の項目）
        if cfg.get("rule_llm", "1") == "1":
            chunks = ds_chunks(item_id)
            llm = pr.LLM(cfg["llm_base_url"], cfg["llm_model"], cfg["llm_api_key"] or None, verbose=False)

            def stop_if_changed():
                it = next((x for x in ds_items() if x["id"] == item_id), None)
                if it is None or it.get("rev", 0) != rev:
                    raise Stale()
            pr.BEFORE_CALL = stop_if_changed
            got: list[dict] = []
            try:
                for i, c in enumerate(chunks, 1):
                    update(task=f"LLM でルールを読み取り中 {i}/{len(chunks)}")
                    got += pr.llm_manual_rules(llm, c)
            except Stale:
                raise
            except BaseException as e:  # LLM に届かないなど（SystemExit を含む）。書式から読み取ったルールは残す
                notes.append(f"LLM でのルール読み取りは途中まで（{str(e)[:80]}）")
            finally:
                pr.BEFORE_CALL = None
            old = ds_rules(item_id)
            on = {pr.rule_key(r): r.get("enabled", True) for r in old}
            rules = pr.dedupe_rules([r for r in old if r["by"] != "LLM"] + got)
            for r in rules:
                r["enabled"] = on.get(pr.rule_key(r), True)
            update(files={f"{item_id}.rules.json": json.dumps(rules, ensure_ascii=False, indent=1)},
                   rules=len(rules), rules_llm=sum(1 for r in rules if r["by"] == "LLM"), llm_done=True)
        update(task="", note=" / ".join(notes))
    except Stale:
        pass
    except Exception as e:
        try:
            update(task="", note=f"処理できませんでした: {str(e)[:120]}")
        except Exception:
            pass
    finally:
        pr.BEFORE_CALL = None
        gc.collect()


def prepare_manual(job: dict, cfg: dict, paras, seg_chars: int):
    """校正方針の渡し方を決める。マニュアルが長ければ、区切りごとに埋め込みで近い項目を top_k 個だけ選んで渡す"""
    chunks = ds_active_chunks()
    job["cites"] = {c["cid"]: (c["name"], c["text"]) for c in chunks}
    if not chunks:
        return skip(job, "retrieve", "マニュアルなし（既定の方針）")
    total = sum(len(c["text"]) for c in chunks)
    mode = cfg.get("manual_mode", "auto")
    if mode == "full" or (mode == "auto" and total <= MANUAL_WARN_CHARS):
        return skip(job, "retrieve", f"全文を使用（{len(chunks)} 項目 / {total:,} 字）")
    k = max(1, int(cfg.get("top_k") or 6))
    with step(job, "retrieve") as st:
        gpu_gate(job)
        try:
            filled = ds_fill_vectors(cfg)
            chunks = ds_active_chunks()
            groups = pr.segments(paras, seg_chars)
            qv = embed(["\n".join(p.text for p in g) for g in groups], cfg, on_progress=pr.PROGRESS)
            order_of = {c["cid"]: i for i, c in enumerate(chunks)}
            picks = []
            for v in qv:
                usable = [c for c in chunks if c.get("vec") and len(c["vec"]) == len(v)]  # 別のモデルで計算したものは使わない
                top = sorted(usable, key=lambda c: -_cos(v, c["vec"]))[:k]
                picks.append(numbered(sorted(top, key=lambda c: order_of[c["cid"]])))
            pr.MANUAL_FOR = lambda gi, picks=picks: picks[gi] if gi < len(picks) else pr.MANUAL
            st["note"] = f"区切り {len(groups)} × 上位 {k} 項目（全 {len(chunks)} 項目）"
            pr.log(f"マニュアルの検索: {st['note']}" + (f"。未計算だった埋め込み {filled} 件を計算" if filled else ""))
        except Cancelled:
            raise
        except Exception as e:
            pr.MANUAL_FOR = None
            st["note"] = "検索できず全文を使用"
            pr.log(f"マニュアルの検索ができないため全文を使います（{str(e)[:200]}）")


def skip(job: dict, key: str, note: str):
    job["steps"][key].update(status="スキップ", note=note)


def worker():
    while True:
        jid = q.get()
        if isinstance(jid, tuple):  # ("ds", マニュアルの id)
            try:
                run_ds_task(jid[1])
            except BaseException:  # ここで止まるとワーカーが死に、以後どのジョブも動かなくなる
                traceback.print_exc()
            q.task_done()
            continue
        if jid not in jobs:
            q.task_done()
            continue
        job = jobs[jid]
        job["status"] = "実行中"
        job["started"] = time.time()

        def log(msg, job=job):
            job["log"].append(time.strftime("%H:%M:%S ") + msg)

        pr.log = log
        pr.BEFORE_CALL = lambda job=job: gpu_gate(job)
        rules_only = job.get("mode") == "rules"
        llm = None
        try:
            src = Path(job["path"])
            with step(job, "extract") as st:
                from docx import Document
                doc = Document(str(src))
                paras = pr.extract_paragraphs(doc)
                if not paras:
                    raise RuntimeError("【書類名】明細書 以降の本文が見つかりません。")
                st["note"] = f"{len(paras)} 段落 / {sum(len(p.text) for p in paras):,} 文字"
                log(f"本文 {st['note']}")
            cfg = load_config()
            apply_config(cfg)
            with step(job, "rules") as st:
                mrules = ds_active_rules()
                by_type = lambda t: [r["from"] for r in mrules if r["type"] == t]
                pr.configure_symbols(prefix=by_type("symbol_prefix"), exclude=by_type("symbol_exclude"))
                ignore_words = by_type("ignore")
                table = pr.symbol_table(paras)
                findings = pr.rule_checks(paras, table)
                got = pr.manual_rule_checks(paras, mrules)
                findings += got
                st["note"] = f"{len(findings)} 件（符号 {len(table)} 個・マニュアルの規則 {len(got)} 件）"
                log(f"ルールチェック（対応表・語の重複・表記ゆれ・マニュアルの規則 {len(mrules)} 個）: {st['note']}")
            pdfs = sorted(x for x in src.parent.glob("*.pdf") if "図" in x.name)
            if not pdfs:
                skip(job, "drawing", "図面PDFなし")
                log("同じフォルダに図面PDF（名前に「図」を含む .pdf）がないため符号照合をスキップ")
            else:
                with step(job, "drawing") as st:
                    cached = pr.ocr_cache_path(pdfs[0]).exists()
                    # OCR は LLM の工程より前にだけ行う。ルールのみなら読み直し、フルなら保存済みの結果があればそれを使う
                    use_ocr = bool(pr.OCR_URL) and (rules_only or not cached) and ocr_running(force=True)
                    log(f"図面PDF: {pdfs[0].name}")
                    got = pr.drawing_symbol_checks(paras, pdfs[0], allow_ocr=use_ocr, table=table)
                    findings += got
                    st["note"] = f"{len(got)} 件" + ("（OCR で読み取り）" if use_ocr else "（保存済みの OCR 結果）" if cached else "")
            if rules_only:
                skip(job, "retrieve", "ルールのみ")
                skip(job, "vote", "ルールのみ")
                skip(job, "verify", "ルールのみ")
                log("ルールのみモード（LLM は使いません）")
            else:
                llm = pr.LLM(cfg["llm_base_url"], cfg["llm_model"], cfg["llm_api_key"] or None, verbose=True)
                seg_chars = int(os.environ.get("PROOF_SEGMENT_CHARS", "1200"))
                prepare_manual(job, cfg, paras, seg_chars)
                with step(job, "vote") as st:
                    gpu_gate(job)
                    adopted, single, gtext = pr.llm_vote(llm, paras, seg_chars)
                    st["note"] = llm.short()
                    log(f"LLM 観点チェック: {llm.summary()}")
                with step(job, "verify") as st:
                    n0 = len(llm.stats)
                    n1 = sum(len(v) for v in single.values())
                    got = pr.llm_verify(llm, single, gtext)
                    adopted += got
                    st["note"] = f"1票の候補 {n1} 件 → {len(got)} 件採用 ・ {llm.short(n0)}" if n1 else "検証する候補なし"
                findings += adopted
            with step(job, "write") as st:
                located = [f for f in findings if pr.locate(f, paras)]
                final = [f for f in pr.dedupe(located) if not pr.is_ignored(f, ignore_words)]
                log(f"指摘 {len(final)} 件（原文に見つからない引用 {len(findings) - len(located)} 件を除外）")
                notes = [f.note for f in final]
                cites = job.get("cites") or {}
                for f in final:  # コメントと CSV には根拠の項目の冒頭も書く（どこに基づく指摘か docx だけで追えるように）
                    used = [c for c in dict.fromkeys(pr.CITE_RE.findall(f.note)) if c[1:-1] in cites]
                    if used:
                        f.note += "".join(f"\n{c} {cites[c[1:-1]][0]}: {cites[c[1:-1]][1][:80]}" for c in used)
                pr.apply_to_docx(doc, final)
                tag = "【AI校正・ルールのみ】" if rules_only else "【AI校正】"
                suffix = "_ルールのみ" if rules_only else ""
                out_docx = src.parent / f"{tag}{src.name}"
                out_csv = src.parent / f"{src.stem}_修正箇所一覧{suffix}.csv"
                out_tbl = src.parent / f"{src.stem}_符号対応表.csv"
                doc.save(str(out_docx))
                import csv
                with open(out_csv, "w", newline="", encoding="utf-8-sig") as fp:
                    w = csv.writer(fp)
                    w.writerow(["No", "段落番号", "修正前", "修正後", "種類", "備考", "要確認", "反映", "検出元", "検出回数"])
                    for i, f in enumerate(final, 1):
                        w.writerow([i, f.label, f.before, f.after, f.kind, f.note, "○" if f.needs_check else "", f.applied, f.source, f"{f.votes}/{pr.VOTE_RUNS}" if f.votes else "ルール"])
                pr.write_table_csv(out_tbl, table)
                job["outputs"] = [str(out_docx), str(out_csv), str(out_tbl)]
                job["rows"] = [(f.label, f.before, f.after, f.kind, n, f.applied) for f, n in zip(final, notes)]
                job["table"] = pr.table_rows(table)
                st["note"] = f"{len(final)} 件"
            job["status"] = "完了"
            if llm:
                log(f"LLM 合計: {llm.summary()}")
        except Cancelled:
            job["status"] = "キャンセル"
            log("キャンセルしました")
        except SystemExit as e:
            job["status"] = "エラー"
            log(str(e))
        except Exception:
            job["status"] = "エラー"
            log(traceback.format_exc()[-1500:])
        for st in job["steps"].values():
            if st["status"] == "待機":
                st["status"] = "—"
        job["current"] = None
        pr.BEFORE_CALL = None
        pr.MANUAL_FOR = None
        pr.configure_symbols()
        job["finished"] = time.time()
        gc.collect()
        q.task_done()


# ---- 画面（左のサイドバー＋本文のレイアウト。一覧の状態は JS で数秒ごとに差し替える） ----

CSS = """
:root{--bg:#f5f6f8;--panel:#fff;--line:#e6e8ec;--text:#1f2329;--sub:#646a73;--faint:#8f959e;
--pri:#3370ff;--pri-h:#245bdb;--pri-bg:#eef3ff;--ok:#00a870;--ok-bg:#e8f8f2;--ng:#e34d59;--ng-bg:#fdecee;
--run:#ed7b2f;--run-bg:#fef3e6;--wait:#646a73;--wait-bg:#f0f1f3}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--text);font:14px/1.6 -apple-system,'Hiragino Sans','Noto Sans JP',sans-serif}
a{color:var(--pri);text-decoration:none}a:hover{text-decoration:underline}
.app{display:flex;min-height:100vh}
.side{width:220px;flex:none;background:var(--panel);border-right:1px solid var(--line);padding:20px 12px;display:flex;flex-direction:column;gap:4px;position:sticky;top:0;height:100vh}
.brand{padding:2px 12px 18px;font-weight:800;font-size:20px;letter-spacing:-.01em;color:var(--pri)}
.nav{display:flex;align-items:center;gap:10px;padding:9px 12px;border-radius:8px;color:var(--text)}
.nav:hover{background:var(--bg);text-decoration:none}.nav.on{background:var(--pri-bg);color:var(--pri);font-weight:600}
.nav svg{width:18px;height:18px;flex:none}
.side .foot{margin-top:auto;padding:12px 8px 0;border-top:1px solid var(--line);font-size:12px;color:var(--faint);word-break:break-all}
.main{flex:1;min-width:0;padding:28px 32px 48px}
.head{display:flex;align-items:flex-start;justify-content:space-between;gap:16px;margin-bottom:20px;flex-wrap:wrap}
.head h1{font-size:22px;margin:0 0 2px}.head p{margin:0;color:var(--sub);font-size:13px}
.crumb{font-size:13px;color:var(--sub);margin-bottom:6px}
.card{background:var(--panel);border:1px solid var(--line);border-radius:12px;margin-bottom:16px;overflow-x:auto}
.card>h2{font-size:15px;margin:0;padding:14px 18px;border-bottom:1px solid var(--line)}
.card>.body{padding:16px 18px}
.bar{display:flex;gap:10px;align-items:center;padding:12px 18px;border-bottom:1px solid var(--line);flex-wrap:wrap}
.search{flex:1;min-width:180px;max-width:320px;padding:7px 12px;border:1px solid var(--line);border-radius:8px;font:inherit}
.btn{display:inline-flex;align-items:center;gap:6px;border:1px solid transparent;border-radius:8px;padding:6px 14px;font:inherit;font-size:13px;cursor:pointer;white-space:nowrap;background:var(--pri);color:#fff}
.btn:hover{background:var(--pri-h);text-decoration:none}.btn:disabled{opacity:.5;cursor:default}
.btn.ghost{background:var(--panel);color:var(--text);border-color:var(--line)}.btn.ghost:hover{border-color:var(--pri);color:var(--pri)}
.btn.sm{padding:3px 10px;font-size:12px}.btn svg{width:15px;height:15px}
.tbl{width:100%;border-collapse:collapse}
.tbl th{font-size:12px;font-weight:600;color:var(--sub);text-align:left;padding:10px 18px;background:#fafbfc;border-bottom:1px solid var(--line)}
.tbl td{padding:12px 18px;border-bottom:1px solid var(--line);vertical-align:middle}
.tbl tr:last-child td{border-bottom:0}.tbl tbody tr:hover{background:#fafbfc}
.tbl tr.new{animation:flash 2.5s}@keyframes flash{from{background:var(--pri-bg)}to{background:transparent}}
.fname{font-weight:600;word-break:break-all}.fdir{font-size:12px;color:var(--faint);word-break:break-all}
.acts{display:flex;gap:6px;justify-content:flex-end;flex-wrap:wrap}
.tag{display:inline-flex;align-items:center;gap:5px;font-size:12px;padding:1px 9px;border-radius:999px;background:var(--wait-bg);color:var(--wait);white-space:nowrap}
.tag.ok{background:var(--ok-bg);color:var(--ok)}.tag.ng{background:var(--ng-bg);color:var(--ng)}
.tag.run{background:var(--run-bg);color:var(--run)}.tag.pri{background:var(--pri-bg);color:var(--pri)}
.tag.run::before{content:"";width:7px;height:7px;border-radius:50%;background:currentColor;animation:blink 1s infinite}
@keyframes blink{50%{opacity:.25}}
.note{font-size:12px;color:var(--faint);margin-top:3px;max-width:340px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.empty{padding:48px 18px;text-align:center;color:var(--faint)}
.muted{color:var(--faint);font-size:12px}
.modal{position:fixed;inset:0;background:rgba(15,20,30,.45);display:none;align-items:center;justify-content:center;padding:16px;z-index:10}
.modal.show{display:flex}
.dlg{background:var(--panel);border-radius:14px;width:100%;max-width:560px;box-shadow:0 12px 40px rgba(0,0,0,.18)}
.dlg h3{margin:0;padding:16px 20px;font-size:16px;border-bottom:1px solid var(--line);display:flex;justify-content:space-between}
.dlg .x{background:none;border:0;font-size:20px;cursor:pointer;color:var(--faint);line-height:1}
.dlg .in{padding:18px 20px}.dlg .ft{padding:12px 20px;border-top:1px solid var(--line);display:flex;justify-content:flex-end;gap:8px}
.drop{border:1.5px dashed #c9cdd4;border-radius:12px;padding:34px 16px;text-align:center;cursor:pointer;background:#fafbfc;transition:.15s}
.drop:hover,.drop.over{border-color:var(--pri);background:var(--pri-bg)}
.drop svg{width:40px;height:40px;color:var(--pri)}.drop b{display:block;margin:8px 0 4px}
.flist{list-style:none;margin:14px 0 0;padding:0;display:flex;flex-direction:column;gap:8px}
.flist li{border:1px solid var(--line);border-radius:8px;padding:8px 12px;font-size:13px}
.flist .row{display:flex;justify-content:space-between;gap:8px}
.prog{height:4px;background:var(--bg);border-radius:4px;margin-top:6px;overflow:hidden}.prog i{display:block;height:100%;width:0;background:var(--pri);transition:width .2s}
.err{color:var(--ng);font-size:13px;margin-top:10px}
.dropall{position:fixed;inset:0;background:rgba(51,112,255,.12);border:3px dashed var(--pri);display:none;align-items:center;justify-content:center;font-size:18px;color:var(--pri);font-weight:600;z-index:20;pointer-events:none}
.dropall.show{display:flex}
.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-bottom:16px}
.stat{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:14px 18px}
.stat .k{font-size:12px;color:var(--sub)}.stat .v{font-size:22px;font-weight:700}
.chips{display:flex;gap:6px;flex-wrap:wrap}
.chip{border:1px solid var(--line);background:var(--panel);border-radius:999px;padding:3px 12px;font-size:12px;cursor:pointer}
.chip.on{background:var(--pri);border-color:var(--pri);color:#fff}
td.diff{min-width:130px}.diff del{background:var(--ng-bg);color:#b42330;text-decoration:line-through}.diff ins{background:var(--ok-bg);color:#00744e;text-decoration:none}
details.log summary{cursor:pointer;padding:14px 18px;font-weight:600;font-size:15px}
pre{margin:0;white-space:pre-wrap;font-size:12px;background:#1f2329;color:#d7dae0;padding:14px 18px;border-radius:0 0 12px 12px;max-height:360px;overflow:auto}
.form{display:grid;grid-template-columns:200px 1fr;gap:16px 20px;align-items:start}
.form label{font-weight:600;font-size:13px;padding-top:7px}
.form input,.form select,textarea{width:100%;padding:7px 10px;border:1px solid var(--line);border-radius:8px;font:inherit;background:var(--panel)}
.form input:focus,.form select:focus,textarea:focus,.search:focus{outline:none;border-color:var(--pri);box-shadow:0 0 0 3px var(--pri-bg)}
.form .hint{font-size:12px;color:var(--faint);margin-top:4px}
.steps{display:grid;grid-template-columns:repeat(7,1fr);gap:0;padding:18px}
.cite{display:inline-block;background:var(--pri-bg);color:var(--pri);border-radius:6px;padding:0 5px;font-size:12px;cursor:help}
.chunk{border:1px solid var(--line);border-radius:8px;padding:8px 12px;font-size:13px;white-space:pre-wrap;margin-bottom:8px}.chunk b{color:var(--pri);margin-right:6px}
.chunks{max-height:60vh;overflow:auto}
.stp{position:relative;padding:0 8px;text-align:center;font-size:12px;color:var(--sub)}
.stp::before{content:"";position:absolute;top:13px;left:-50%;right:50%;height:2px;background:var(--line)}
.stp:first-child::before{display:none}.stp.done::before,.stp.run::before{background:var(--pri)}
.stp .dot{position:relative;z-index:1;width:28px;height:28px;margin:0 auto 6px;border-radius:50%;display:grid;place-items:center;background:var(--wait-bg);color:var(--sub);font-weight:700;font-size:13px}
.stp.done .dot{background:var(--pri);color:#fff}.stp.run .dot{background:var(--run);color:#fff;animation:blink 1.2s infinite}
.stp.ng .dot{background:var(--ng);color:#fff}.stp.skip .dot{background:var(--bg);color:var(--faint)}
.stp b{display:block;color:var(--text);font-size:13px}
.stp .pct{font-size:15px;font-weight:700;color:var(--run);margin-top:2px}
.pbar{height:4px;background:var(--bg);border-radius:4px;margin:4px 10px 0;overflow:hidden}.pbar i{display:block;height:100%;background:var(--run);transition:width .4s}.stp .res{font-size:11px;color:var(--faint)}
.stp .sn{font-size:11px;color:var(--sub);margin-top:2px;word-break:break-all}
.switch{position:relative;width:36px;height:20px;flex:none;display:inline-block}.switch input{display:none}
.switch i{position:absolute;inset:0;border-radius:999px;background:#c9cdd4;transition:.15s;cursor:pointer}
.switch i::after{content:"";position:absolute;top:2px;left:2px;width:16px;height:16px;border-radius:50%;background:#fff;transition:.15s}
.switch input:checked+i{background:var(--pri)}.switch input:checked+i::after{left:18px}
.btn.danger{background:var(--panel);color:var(--ng);border-color:var(--line)}.btn.danger:hover{border-color:var(--ng);background:var(--ng-bg)}
.btn.red{background:var(--ng)}.btn.red:hover{background:#c73a46}
.dlg.wide{max-width:820px}.dlg textarea{min-height:360px;font-size:13px;line-height:1.7}
.warn{color:#8a4a12;background:var(--run-bg);border-radius:8px;padding:8px 12px;font-size:13px;margin-top:10px}
details.def summary{cursor:pointer;color:var(--sub);font-size:13px}details.def pre{border-radius:8px;margin-top:8px;background:#fafbfc;color:var(--text);border:1px solid var(--line)}
.toast{position:fixed;right:20px;bottom:20px;background:#1f2329;color:#fff;padding:10px 16px;border-radius:8px;font-size:13px;opacity:0;transition:.2s;z-index:30}
.toast.show{opacity:1}
@media (max-width:760px){
 .app{flex-direction:column}.side{width:auto;height:auto;position:static;flex-direction:row;align-items:center;padding:10px 12px;border-right:0;border-bottom:1px solid var(--line);overflow-x:auto}
 .brand{padding:0 10px 0 2px;font-size:17px}.side .foot{display:none}.nav{white-space:nowrap;flex:none}.main{width:100%;padding:18px 16px 40px}
 .stats{grid-template-columns:repeat(3,1fr);gap:8px}.stat{padding:10px 12px}.stat .v{font-size:18px}.bar .muted{display:none}
 .tbl .hide-sm{display:none}.steps{grid-template-columns:repeat(3,1fr);row-gap:16px}.stp:nth-child(3n+1)::before{display:none}.tbl td,.tbl th{padding:10px 12px}.form{grid-template-columns:1fr}.form label{padding-top:0}
}
"""

ICONS = {
    "doc": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"/><path d="M14 3v6h6M8 13h8M8 17h5"/></svg>',
    "db": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><ellipse cx="12" cy="5.5" rx="8" ry="3"/><path d="M4 5.5v13c0 1.7 3.6 3 8 3s8-1.3 8-3v-13M4 12c0 1.7 3.6 3 8 3s8-1.3 8-3"/></svg>',
    "hist": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></svg>',
    "gear": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z"/></svg>',
    "up": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M12 16V4M7 9l5-5 5 5M4 20h16"/></svg>',
    "down": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M12 4v12M7 11l5 5 5-5M4 20h16"/></svg>',
    "play": '<svg viewBox="0 0 24 24" fill="currentColor"><path d="M8 5v14l11-7z"/></svg>',
}

STATUS_TAG = {"待機中": "", "実行中": "run", "完了": "ok", "エラー": "ng", "キャンセル": ""}
STEP_NAME = {k: (i, name, res) for i, (k, name, res) in enumerate(STEPS, 1)}
APPLIED = {"change": ("変更履歴", "ok"), "comment": ("コメント", "pri"), "report-only": ("一覧のみ", "")}


def shell(title: str, active: str, body: str, script: str = "", busy: bool = False) -> str:
    nav = "".join(f'<a class="nav{" on" if key == active else ""}" href="{href}">{ICONS[icon]}<span>{label}</span></a>'
                  for key, href, icon, label in (("docs", "/", "doc", "明細書"), ("dataset", "/dataset", "db", "データセット"), ("jobs", "/jobs", "hist", "実行履歴"), ("settings", "/settings", "gear", "設定")))
    return f"""<!doctype html><html lang=ja><meta charset=utf-8><meta name=viewport content="width=device-width,initial-scale=1">
<title>{html.escape(title)} - Proofy</title><style>{CSS}</style>
<body data-busy={int(busy)}><div class=app><aside class=side><div class=brand>Proofy</div>{nav}
<div class=foot>LLM<br>{html.escape(load_config()["llm_model"])}</div></aside>
<main class=main id=main>{body}</main></div><div class=toast id=toast></div>{script}</body></html>"""


def rel(p) -> str:
    return str(Path(p).relative_to(ROOT))


def fmt_time(t: float) -> str:
    return time.strftime("%Y-%m-%d %H:%M", time.localtime(t))


def fmt_dur(sec: float) -> str:
    sec = int(sec)
    return f"{sec // 60}分{sec % 60:02d}秒" if sec >= 60 else f"{sec}秒"


def result_files(src: Path) -> list[tuple[str, Path]]:
    """明細書の横にある校正結果（フル → ルールのみ の順）"""
    cands = [("校正結果 docx", src.parent / f"【AI校正】{src.name}"), ("修正箇所 CSV", src.parent / f"{src.stem}_修正箇所一覧.csv"),
             ("符号対応表 CSV", src.parent / f"{src.stem}_符号対応表.csv"),
             ("ルールのみ docx", src.parent / f"【AI校正・ルールのみ】{src.name}"), ("ルールのみ CSV", src.parent / f"{src.stem}_修正箇所一覧_ルールのみ.csv")]
    return [(k, p) for k, p in cands if p.exists()]


def dl_url(p) -> str:
    return "/download?path=" + urllib.parse.quote(rel(p))


def latest_jobs() -> dict[str, str]:
    out = {}
    for jid in order:
        out[jobs[jid]["path"]] = jid
    return out


# 全体の % を見積もるときの工程ごとの重み（2万字級のフル校正での所要時間のおおよその割合）。スキップした工程は数えない
STEP_WEIGHT = {"extract": 1, "rules": 3, "drawing": 8, "retrieve": 3, "vote": 75, "verify": 8, "write": 2}


def overall_pct(j: dict) -> int:
    steps = j.get("steps") or {}
    use = {k: w for k, w in STEP_WEIGHT.items() if steps.get(k, {}).get("status") != "スキップ"}
    total = sum(use.values()) or 1
    done = sum(w * (100 if steps.get(k, {}).get("status") == "完了" else steps.get(k, {}).get("pct", 0)) for k, w in use.items())
    return min(100, int(done / total))


def job_tag(jid: str, show_mode: bool = True) -> str:
    j = jobs[jid]
    st = j["status"]
    label = st
    if st == "待機中":
        ahead = sum(1 for o in order if o < jid and jobs[o]["status"] in ("待機中", "実行中"))
        label = f"待機中（{ahead}件待ち）" if ahead else st
    elif st == "実行中" and j.get("started"):
        cur = STEP_NAME.get(j.get("current") or "")
        pct = (j.get("steps") or {}).get(j.get("current") or "", {}).get("pct", 0)
        label = (f"{cur[0]}/{len(STEPS)} {cur[1]} {pct}% ・ 全体 {overall_pct(j)}% ・ {fmt_dur(time.time() - j['started'])}"
                 if cur else f"実行中 {fmt_dur(time.time() - j['started'])}")
    mode = "ルールのみ" if show_mode and j.get("mode") == "rules" else ""
    return f'<span class="tag {STATUS_TAG[st]}">{html.escape(label)}</span>' + (f' <span class=muted>{mode}</span>' if mode else "")


def drawing_tag(src: Path) -> str:
    pdfs = sorted(p for p in src.parent.glob("*.pdf") if "図" in p.name)
    if not pdfs:
        return '<span class=muted>なし</span>'
    if pr.ocr_cache_path(pdfs[0]).exists():
        return f'<span class="tag ok" title="{html.escape(pdfs[0].name)}">読み取り済み</span>'
    return f'<span class=tag title="{html.escape(pdfs[0].name)}">未読み取り</span>'


def doc_rows() -> str:
    latest = latest_jobs()
    rows = []
    for p in list_docx()[:200]:
        r = rel(p)
        jid = latest.get(str(p))
        busy = bool(jid and jobs[jid]["status"] in ("待機中", "実行中"))
        results = result_files(p)
        if jid:
            state = job_tag(jid)
            if jobs[jid]["status"] == "実行中" and jobs[jid]["log"]:
                state += f'<div class=note title="{html.escape(jobs[jid]["log"][-1])}">{html.escape(jobs[jid]["log"][-1][9:])}</div>'
        elif results:
            state = '<span class="tag ok">校正済み</span>'
        else:
            state = '<span class=tag>未校正</span>'
        acts = []
        if jid:
            acts.append(f'<a class="btn ghost sm" href="/job/{jid}">詳細</a>')
        if results:
            acts.append(f'<a class="btn ghost sm" href="{dl_url(results[0][1])}" title="{html.escape(results[0][1].name)}">{ICONS["down"]}結果</a>')
        if busy:
            acts.append(f'<button class="btn danger sm" data-cancel="{jid}">中止</button>')
        dis = " disabled" if busy else ""
        acts.append(f'<button class="btn ghost sm" data-run="{html.escape(r)}" data-mode=rules{dis}>ルールのみ</button>')
        acts.append(f'<button class="btn sm" data-run="{html.escape(r)}" data-mode=full{dis}>{ICONS["play"]}校正</button>')
        folder = str(Path(r).parent)
        rows.append(f'<tr data-name="{html.escape(r.lower())}"><td><div class=fname>{html.escape(p.name)}</div>'
                    f'<div class=fdir>{html.escape("Documents/" + folder if folder != "." else "Documents")}</div></td>'
                    f'<td class=hide-sm>{drawing_tag(p)}</td><td class="hide-sm muted">{fmt_time(p.stat().st_mtime)}</td>'
                    f'<td>{state}</td><td><div class=acts>{"".join(acts)}</div></td></tr>')
    return "".join(rows) or '<tr><td colspan=5 class=empty>明細書がまだありません。「ファイルを追加」から .docx をアップロードしてください。</td></tr>'


def any_busy() -> bool:
    return any(jobs[j]["status"] in ("待機中", "実行中") for j in order)


INDEX_JS = r"""<script>
const $ = s => document.querySelector(s);
const toast = t => { const e = $('#toast'); e.textContent = t; e.classList.add('show'); clearTimeout(e._t); e._t = setTimeout(() => e.classList.remove('show'), 3000); };
const modal = $('#upModal'), pick = $('#pick'), drop = $('#drop'), flist = $('#flist'), upErr = $('#upErr');
let fresh = null;

// 一覧の検索
$('#q').oninput = filter;
function filter() { const v = $('#q').value.trim().toLowerCase(); document.querySelectorAll('#rows tr[data-name]').forEach(tr => tr.hidden = v && !tr.dataset.name.includes(v)); }

// 一覧の更新（実行中は数秒ごと）
async function refresh() {
  let d;
  try { const r = await fetch('/api/rows'); if (!r.ok) throw 0; d = await r.json(); } catch (_) { clearTimeout(window._poll); window._poll = setTimeout(refresh, 4000); return; }
  $('#rows').innerHTML = d.rows; $('#stats').innerHTML = d.stats; filter();
  if (fresh) { const tr = [...document.querySelectorAll('#rows tr[data-name]')].find(t => t.dataset.name === fresh); if (tr) tr.classList.add('new'); fresh = null; }
  clearTimeout(window._poll); if (d.busy) window._poll = setTimeout(refresh, 4000);
}
if (document.body.dataset.busy === '1') window._poll = setTimeout(refresh, 4000);

// 中止ボタン
document.addEventListener('click', async e => {
  const b = e.target.closest('[data-cancel]'); if (!b) return;
  b.disabled = true; await fetch('/cancel/' + b.dataset.cancel, {method: 'POST'}); toast('中止しています（実行中の LLM 呼び出しが終わりしだい止まります）'); refresh();
});
// 実行ボタン
let pending = null;
document.addEventListener('click', e => {
  const b = e.target.closest('[data-run]'); if (!b) return;
  pending = {path: b.dataset.run, mode: b.dataset.mode};
  if (pending.mode === 'full') { $('#runName').textContent = pending.path; $('#runModal').classList.add('show'); } else run();
});
$('#runOk').onclick = () => { $('#runModal').classList.remove('show'); run(); };
async function run() {
  const body = new URLSearchParams({path: pending.path, mode: pending.mode});
  const r = await fetch('/run', {method: 'POST', body, headers: {'X-Requested-With': 'fetch'}});
  if (!r.ok) { toast('開始できませんでした'); return; }
  toast(pending.mode === 'rules' ? 'ルールのみの校正を開始しました' : '校正を開始しました（完了まで画面を閉じても大丈夫です）');
  refresh();
}

// モーダル
document.querySelectorAll('[data-open]').forEach(b => b.onclick = () => { reset(); modal.classList.add('show'); });
document.querySelectorAll('[data-close]').forEach(b => b.onclick = () => b.closest('.modal').classList.remove('show'));
document.querySelectorAll('.modal').forEach(m => m.addEventListener('click', e => { if (e.target === m && !m.dataset.lock) m.classList.remove('show'); }));
function reset() { flist.innerHTML = ''; upErr.textContent = ''; pick.value = ''; delete modal.dataset.lock; }

// ファイル選択・ドラッグ＆ドロップ（画面のどこに落としても受け付ける）
drop.onclick = () => pick.click();
pick.onchange = () => upload([...pick.files]);
let depth = 0;
addEventListener('dragenter', e => { if (!e.dataTransfer.types.includes('Files')) return; e.preventDefault(); depth++; $('#dropAll').classList.add('show'); drop.classList.add('over'); });
addEventListener('dragleave', e => { if (--depth <= 0) { depth = 0; $('#dropAll').classList.remove('show'); drop.classList.remove('over'); } });
addEventListener('dragover', e => e.preventDefault());
addEventListener('drop', e => { e.preventDefault(); depth = 0; $('#dropAll').classList.remove('show'); drop.classList.remove('over');
  if (e.dataTransfer.files.length) { if (!modal.classList.contains('show')) { reset(); modal.classList.add('show'); } upload([...e.dataTransfer.files]); } });

function sendFile(f, folder, bar) {
  return new Promise(res => {
    const x = new XMLHttpRequest();
    x.open('POST', `/upload?folder=${encodeURIComponent(folder)}&name=${encodeURIComponent(f.name)}`);
    x.upload.onprogress = e => { if (e.lengthComputable) bar.style.width = (e.loaded / e.total * 100) + '%'; };
    x.onload = () => res(x.status === 200 ? '' : x.responseText);
    x.onerror = () => res('通信エラー');
    x.send(f);
  });
}
async function upload(files) {
  upErr.textContent = '';
  const skipped = files.filter(f => !/\.(docx|pdf)$/i.test(f.name) || f.name.startsWith('~$'));
  files = files.filter(f => !skipped.includes(f));
  const docx = files.find(f => /\.docx$/i.test(f.name));
  if (!docx) { upErr.textContent = '明細書の .docx を含めてください（図面PDFは .docx と一緒に選びます）。'; return; }
  const folder = docx.name.replace(/\.docx$/i, '');
  modal.dataset.lock = '1';
  flist.innerHTML = files.map((f, i) => `<li><div class=row><span>${f.name.replace(/</g, '&lt;')}</span><span class=muted id=s${i}>${(f.size / 1024 / 1024).toFixed(1)} MB</span></div><div class=prog><i id=p${i}></i></div></li>`).join('');
  let ok = true;
  for (const [i, f] of files.entries()) {
    const err = await sendFile(f, folder, $('#p' + i));
    $('#s' + i).textContent = err ? '失敗' : '完了';
    if (err) { ok = false; upErr.textContent = `${f.name}: ${err}`; }
  }
  delete modal.dataset.lock;
  if (skipped.length) upErr.textContent += ` 対象外のため除外: ${skipped.map(f => f.name).join('、')}`;
  if (ok) {
    fresh = (`${$('#upDir').value}/${folder}/${docx.name}`).toLowerCase();
    setTimeout(() => modal.classList.remove('show'), skipped.length ? 2500 : 600);
    toast(`${docx.name} を追加しました`);
    refresh();
  }
}
</script>"""


def doc_stats() -> str:
    docs = list_docx()
    done = sum(1 for p in docs if result_files(p))
    running = sum(1 for j in order if jobs[j]["status"] in ("待機中", "実行中"))
    return (f'<div class=stat><div class=k>明細書</div><div class=v>{len(docs)}</div></div>'
            f'<div class=stat><div class=k>校正済み</div><div class=v>{done}</div></div>'
            f'<div class=stat><div class=k>実行中・待機中</div><div class=v>{running}</div></div>')


def page_index() -> str:
    body = f"""<div class=head><div><h1>明細書</h1></div>
<button class=btn data-open>{ICONS["up"]}ファイルを追加</button></div>
<div class=stats id=stats>{doc_stats()}</div>
<div class=card><div class=bar><input class=search id=q placeholder="ファイル名で検索"></div>
<table class=tbl><thead><tr><th>ファイル</th><th class=hide-sm>図面PDF</th><th class=hide-sm>更新日時</th><th>状態</th><th style=text-align:right>操作</th></tr></thead>
<tbody id=rows>{doc_rows()}</tbody></table></div>

<div class=modal id=upModal><div class=dlg><h3>ファイルを追加<button class=x data-close>×</button></h3><div class=in>
<div class=drop id=drop>{ICONS["up"]}<b>クリックして選択、またはドラッグ＆ドロップ</b><span class=muted>明細書（.docx）と図面PDF（名前に「図」を含む .pdf）をまとめて選べます</span></div>
<input type=file id=pick multiple accept=".docx,.pdf" hidden><input type=hidden id=upDir value="{html.escape(UPLOAD_DIR)}">
<ul class=flist id=flist></ul><div class=err id=upErr></div>
<p class=muted style=margin-bottom:0>保存先: Documents/{html.escape(UPLOAD_DIR)}/&lt;明細書名&gt;/ （同じ名前のファイルは上書きします）</p></div></div></div>

<div class=modal id=runModal><div class=dlg><h3>校正を開始<button class=x data-close>×</button></h3><div class=in>
<p style=margin-top:0><b id=runName></b></p>
<p>2万字で 30分前後かかります。</p></div>
<div class=ft><button class="btn ghost" data-close>キャンセル</button><button class=btn id=runOk>{ICONS["play"]}開始する</button></div></div></div>
<div class=dropall id=dropAll>ここにドロップして追加</div>"""
    return shell("明細書", "docs", body, INDEX_JS, any_busy())


def page_jobs() -> str:
    rows = []
    for jid in reversed(order):
        j = jobs[jid]
        el = fmt_dur((j.get("finished") or time.time()) - j["started"]) if j.get("started") else ""
        rows.append(f'<tr><td><a class=fname href="/job/{jid}">{html.escape(Path(j["path"]).name)}</a><div class=fdir>{html.escape(rel(Path(j["path"]).parent))}</div></td>'
                    f'<td>{"ルールのみ" if j.get("mode") == "rules" else "フル"}</td><td>{job_tag(jid)}</td>'
                    f'<td class="hide-sm muted">{fmt_time(int(jid, 16) / 1000)}</td><td class="hide-sm muted">{el}</td>'
                    f'<td class=hide-sm>{len(j.get("rows", [])) if j["status"] == "完了" else ""}</td></tr>')
    table = "".join(rows) or '<tr><td colspan=6 class=empty>まだ実行していません。アプリを再起動すると履歴は消えます（結果ファイルは残ります）。</td></tr>'
    body = f"""<div class=head><div><h1>実行履歴</h1></div></div>
<div class=card><table class=tbl><thead><tr><th>ファイル</th><th>モード</th><th>状態</th><th class=hide-sm>開始</th><th class=hide-sm>所要時間</th><th class=hide-sm>指摘</th></tr></thead><tbody>{table}</tbody></table></div>"""
    script = "<script>setTimeout(() => location.reload(), 5000)</script>" if any_busy() else ""
    return shell("実行履歴", "jobs", body, script)


def page_settings(msg: str = "") -> str:
    cfg = load_config()
    st = check_status(cfg)
    models = list_models(cfg)
    opts = "".join(f"<option value=\"{html.escape(m)}\" {'selected' if m == cfg['llm_model'] else ''}>{html.escape(m)}</option>" for m in models)
    if cfg["llm_model"] not in models:
        opts = f"<option value=\"{html.escape(cfg['llm_model'])}\" selected>{html.escape(cfg['llm_model'])}（見つかりません）</option>" + opts

    def row(label, inner, note=""):
        return f"<label>{label}</label><div>{inner}<div class=hint>{note}</div></div>"

    def inp(name, val, typ="text"):
        return f"<input name={name} type={typ} value=\"{html.escape(str(val))}\">"

    sel = lambda name, val, choices: f"<select name={name}>" + "".join(f"<option {'selected' if c == val else ''}>{c}</option>" for c in choices) + "</select>"
    sel_kv = lambda name, val, choices: f"<select name={name}>" + "".join(f"<option value=\"{html.escape(k)}\" {'selected' if k == val else ''}>{v}</option>" for k, v in choices) + "</select>"
    badge = lambda k: f"<span class='tag {'ok' if st[k][0] == 'ok' else 'ng' if st[k][0] == 'ng' else ''}'>{html.escape(st[k][1])}</span>"
    toast = f"<script>const t=document.getElementById('toast');t.textContent={json.dumps(msg)};t.classList.add('show');setTimeout(()=>t.classList.remove('show'),3000)</script>" if msg else ""
    body = f"""<div class=head><div><h1>設定</h1></div></div>
<div class=card><h2>接続状態</h2><div class="body form">{row("LLM", badge("llm"))}{row("図面OCR", badge("ocr"))}</div></div>
<form method=post action=/settings>
<div class=card><h2>LLM</h2><div class="body form">
{row("LLM の URL", inp("llm_base_url", cfg["llm_base_url"]), "OpenAI 互換 API の URL（例: Olares Router の共有エントランス http://xxxx.shared.olares.com/v1）")}
{row("モデル", f"<select name=llm_model>{opts}</select>", "上の URL から取得したモデル一覧。日本語に強い 27B 以上を推奨")}
{row("API キー", inp("llm_api_key", cfg["llm_api_key"], "password"), "Olares 内の LLM なら空欄でかまいません")}
{row("推論の深さ", sel("reasoning", cfg["reasoning"], ["low", "medium", "xhigh", ""]), "low が速く、xhigh は精度が上がる代わりに数倍遅い")}
{row("多数決の回数", sel("vote_runs", str(cfg["vote_runs"]), ["1", "3", "5"]), "3 推奨。全回一致の指摘だけ変更履歴にします")}
{row("同時実行数", sel("parallel", str(cfg["parallel"]), ["1", "2", "3", "4"]), "LLM 側の同時実行上限に合わせる")}
</div></div>
<div class=card><h2>図面OCR（任意）</h2><div class="body form">
{row("OCR の URL", inp("ocr_url", cfg["ocr_url"]))}
</div></div>
<div class=card><h2>処理の内容</h2><div class="body form">
{row("全文の観点チェック（LLM）", sel_kv("whole_views", cfg.get("whole_views", "0"), [("0", "使わない（推奨）"), ("1", "使う")]))}
</div></div>
<div class=card><h2>校正方針（<a href=/dataset>データセット</a>）</h2><div class="body form">
{row("マニュアルの渡し方", sel_kv("manual_mode", cfg.get("manual_mode", "auto"), [("auto", f"自動（{MANUAL_WARN_CHARS:,} 字までは全文、超えたら検索）"), ("full", "常に全文"), ("search", "常に検索")]))}
{row("埋め込みモデル", sel_kv("embed_model", cfg.get("embed_model", ""), [("", f"自動（{html.escape(embed_model(dict(cfg, embed_model=''), models) or '見つかりません')}）")] + [(html.escape(m), html.escape(m)) for m in models if "embed" in m.lower()]))}
{row("検索で渡す項目数", sel("top_k", str(cfg.get("top_k", "6")), ["4", "6", "8", "12"]))}
{row("ルールの読み取り（LLM）", sel_kv("rule_llm", cfg.get("rule_llm", "1"), [("1", "使う"), ("0", "使わない（書式から読み取れるものだけ）")]))}
</div></div>
<button class=btn>保存して接続確認</button></form>"""
    return shell("設定", "settings", body, toast)


JOB_JS = r"""<script>
// 中止・対応表の絞り込み・種類で絞り込み
document.addEventListener('click', async e => {
  const x = e.target.closest('#cancel');
  if (x) { x.disabled = true; x.textContent = '中止しています…'; await fetch('/cancel/' + x.dataset.id, {method: 'POST'}); return; }
  const t = e.target.closest('[data-t]');
  if (t) { t.parentNode.querySelectorAll('.chip').forEach(c => c.classList.toggle('on', c === t));
    document.querySelectorAll('#ttab tr').forEach(tr => tr.hidden = t.dataset.t === 'bad' && tr.dataset.bad !== '1'); return; }
  const c = e.target.closest('.chip[data-k]'); if (!c) return;
  c.parentNode.querySelectorAll('.chip').forEach(x => x.classList.toggle('on', x === c));
  document.querySelectorAll('#finds tr[data-kind]').forEach(tr => tr.hidden = c.dataset.k && tr.dataset.kind !== c.dataset.k);
});
// 実行中は本文だけ差し替える
async function poll() {
  const r = await fetch(location.href); if (!r.ok) return;
  const doc = new DOMParser().parseFromString(await r.text(), 'text/html');
  const log = document.querySelector('details.log'), open = log ? log.open : true;
  document.getElementById('main').innerHTML = doc.getElementById('main').innerHTML;
  const nl = document.querySelector('details.log'); if (nl) { nl.open = open; const pre = nl.querySelector('pre'); pre.scrollTop = pre.scrollHeight; }
  if (doc.body.dataset.busy === '1') setTimeout(poll, 4000);
}
const pre = document.querySelector('details.log pre'); if (pre) pre.scrollTop = pre.scrollHeight;
if (document.body.dataset.busy === '1') setTimeout(poll, 4000);
</script>"""


def cite_html(note: str, cites: dict) -> str:
    """備考の［M1-2］を、マニュアルの該当項目をツールチップで見られる印にする"""
    def rep(m):
        c = cites.get(m.group(0)[1:-1])
        return f'<span class=cite title="{html.escape(c[0] + chr(10) + c[1])}">{m.group(0)}</span>' if c else m.group(0)
    return pr.CITE_RE.sub(rep, html.escape(note))


def page_job(jid: str) -> str:
    j = jobs[jid]
    src = Path(j["path"])
    busy = j["status"] in ("待機中", "実行中")
    el = fmt_dur((j.get("finished") or time.time()) - j["started"]) if j.get("started") else "—"
    outs = [Path(o) for o in j.get("outputs", []) if Path(o).exists()]
    dl = "".join(f'<a class="btn{"" if i == 0 else " ghost"}" href="{dl_url(o)}">{ICONS["down"]}{"docx" if o.suffix == ".docx" else "対応表 CSV" if "符号対応表" in o.name else "修正箇所 CSV"}</a>'
                 for i, o in enumerate(outs))
    if busy:
        dl += f'<button class="btn danger" id=cancel data-id="{jid}">中止</button>'
    cls = {"完了": "done", "実行中": "run", "エラー": "ng", "キャンセル": "ng", "スキップ": "skip"}
    steps = j.get("steps") or {}
    stepper = "".join(
        f'<div class="stp {cls.get(steps.get(k, {}).get("status"), "")}"><div class=dot>{"✓" if steps.get(k, {}).get("status") == "完了" else i}</div>'
        f'<b>{name}</b><div class=res>{fmt_dur(steps[k]["secs"]) if steps.get(k, {}).get("secs") is not None else "&nbsp;"}</div>'
        + (f'<div class=pct>{steps[k].get("pct", 0)}%</div><div class=pbar><i style="width:{steps[k].get("pct", 0)}%"></i></div>'
           if steps.get(k, {}).get("status") == "実行中" else "")
        + f'<div class=sn>{html.escape(steps.get(k, {}).get("note", "") if steps.get(k, {}).get("status") != "待機" else "")}</div></div>'
        for i, (k, name, res) in enumerate(STEPS, 1))
    trows = j.get("table") or []
    ttable = ""
    if trows:
        bad = sum(1 for r in trows if r[5])
        ttr = "".join(f'<tr data-bad="{1 if r[5] else 0}"><td style=white-space:nowrap><b>{html.escape(r[0])}</b></td><td>{html.escape(r[1])}</td>'
                      f'<td>{html.escape(r[2]) or "<span class=muted>—</span>"}</td><td class="hide-sm muted">{html.escape(r[3])}</td>'
                      f'<td class=hide-sm>{r[4]}</td><td>{"".join(f"<span class='tag ng'>{html.escape(x)}</span> " for x in r[5].split("、") if x)}</td></tr>' for r in trows)
        ttable = f"""<div class=card><h2>名称と符号の対応表</h2><div class=bar><div class=chips><button class="chip on" data-t=all>すべて {len(trows)}</button><button class=chip data-t=bad>矛盾あり {bad}</button></div></div>
<table class=tbl><thead><tr><th>符号</th><th>本文の名称（回数）</th><th>符号の説明</th><th class=hide-sm>初出</th><th class=hide-sm>図面</th><th>矛盾</th></tr></thead><tbody id=ttab>{ttr}</tbody></table></div>"""
    rows = j.get("rows", [])
    kinds: dict[str, int] = {}
    for r in rows:
        kinds[r[3]] = kinds.get(r[3], 0) + 1
    applied = {k: sum(1 for r in rows if r[5] == k) for k in APPLIED}
    stats = ""
    if j["status"] == "完了":
        stats = f"""<div class=stats><div class=stat><div class=k>指摘</div><div class=v>{len(rows)}</div></div>
<div class=stat><div class=k>変更履歴で反映</div><div class=v>{applied["change"]}</div></div>
<div class=stat><div class=k>コメント（要確認）</div><div class=v>{applied["comment"]}</div></div>
<div class=stat><div class=k>所要時間</div><div class=v style=font-size:18px>{el}</div></div></div>"""
    table = ""
    if rows:
        chips = '<button class="chip on" data-k="">すべて ' + str(len(rows)) + "</button>" + "".join(
            f'<button class=chip data-k="{html.escape(k)}">{html.escape(k)} {n}</button>' for k, n in sorted(kinds.items(), key=lambda x: -x[1]))
        trs = "".join(
            f'<tr data-kind="{html.escape(kind)}"><td class="muted" style=white-space:nowrap>{html.escape(label)}</td>'
            + (f'<td class=diff><del>{html.escape(before)}</del></td><td class=diff><ins>{html.escape(after)}</ins></td>' if after else
               f'<td class=diff>{html.escape(before)}</td><td class=muted>—</td>')
            + f'<td><span class=tag>{html.escape(kind)}</span></td><td class="hide-sm" style=font-size:13px>{cite_html(note, j.get("cites") or {})}</td>'
            f'<td><span class="tag {APPLIED.get(ap, (ap, ""))[1]}">{html.escape(APPLIED.get(ap, (ap, ""))[0])}</span></td></tr>'
            for label, before, after, kind, note, ap in rows)
        table = f"""<div class=card><h2>修正箇所</h2><div class=bar><div class=chips>{chips}</div></div>
<table class=tbl><thead><tr><th>段落</th><th>修正前</th><th>修正後</th><th>種類</th><th class=hide-sm>備考</th><th>反映</th></tr></thead><tbody id=finds>{trs}</tbody></table></div>"""
    body = f"""<div class=crumb><a href=/>明細書</a> / {html.escape(src.name)}</div>
<div class=head><div><h1>{html.escape(src.name)}</h1><p>{job_tag(jid, show_mode=False)} <span class=muted>{"ルールのみ" if j.get("mode") == "rules" else "フル校正"} ・ 経過 {el} ・ Documents/{html.escape(rel(src.parent))}</span></p></div>
<div class=acts>{dl}</div></div>
<div class=card><h2>工程</h2><div class=steps>{stepper}</div></div>
{stats}{table}{ttable}
<div class=card><details class=log{" open" if busy or j["status"] == "エラー" else ""}><summary>ログ</summary><pre>{html.escape(chr(10).join(j["log"][-200:])) or "（待機中）"}</pre></details></div>"""
    return shell(src.name, "docs", body, JOB_JS, busy)


def ds_rows() -> str:
    rows = []
    for x in ds_items():
        on = " checked" if x.get("enabled") else ""
        emb = ('<span class="tag ok">埋め込み済み</span>' if x.get("embedded") else '<span class=tag>埋め込み未計算</span>') if x.get("chunks") else ""
        img = ""
        if x.get("images"):
            unread = x.get("images_unread", 0)
            img = f'<span class="tag {"run" if unread else "ok"}">画像 {x["images"]}{f"（未読み取り {unread}）" if unread else "（読み取り済み）"}</span>'
        note = f'<div class=note title="{html.escape(x.get("note", ""))}">{html.escape(x.get("note", ""))}</div>' if x.get("note") else ""
        if x.get("task"):
            note = f'<div><span class="tag run">{html.escape(x["task"])}</span></div>' + note
        rules = f'<span class=tag>ルール {x.get("rules", 0)}（LLM {x.get("rules_llm", 0)}）</span>' if x.get("rules") else ""
        rows.append(f'<tr data-name="{html.escape(x["name"].lower())}" data-task="{1 if x.get("task") else 0}"><td><div class=fname>{html.escape(x["name"])}</div>'
                    f'<div class=fdir>{html.escape(x.get("code", ""))} ・ {html.escape(x.get("source", ""))}</div></td>'
                    f'<td class=hide-sm>{x.get("chars", 0):,} 字<div class=muted>{x.get("chunks", 0)} チャンク</div></td>'
                    f'<td class=hide-sm>{emb} {rules} {img}{note}</td><td class="hide-sm muted">{fmt_time(x.get("updated", 0))}</td>'
                    f'<td><label class=switch title="校正に使う"><input type=checkbox data-toggle="{x["id"]}"{on}><i></i></label></td>'
                    f'<td><div class=acts><button class="btn ghost sm" data-chunks="{x["id"]}">チャンク</button>'
                    f'<button class="btn ghost sm" data-rules="{x["id"]}">ルール</button>'
                    f'<button class="btn ghost sm" data-edit="{x["id"]}">編集</button>'
                    f'<button class="btn ghost sm" data-re="{x["id"]}" title="画像の読み取り・チャンク分割・埋め込みをやり直す">再処理</button>'
                    f'<a class="btn ghost sm" href="/dataset/{x["id"]}/download" title="取り出した本文をダウンロード">{ICONS["down"]}</a>'
                    f'<button class="btn danger sm" data-del="{x["id"]}" data-name="{html.escape(x["name"])}">削除</button></div></td></tr>')
    return "".join(rows) or '<tr><td colspan=6 class=empty>マニュアルがまだありません。「マニュアルを追加」から取り込んでください。<br>マニュアルがない間は下の既定の方針で校正します。</td></tr>'


def ds_summary() -> str:
    items = ds_items()
    on = [x for x in items if x.get("enabled")]
    chars = sum(x.get("chars", 0) for x in on)
    return (f'<div class=stat><div class=k>マニュアル</div><div class=v>{len(items)}</div></div>'
            f'<div class=stat><div class=k>校正に使用中</div><div class=v>{len(on)}</div></div>'
            f'<div class=stat><div class=k>使用中のチャンク</div><div class=v>{sum(x.get("chunks", 0) for x in on)}</div></div>'
            f'<div class=stat><div class=k>使用中の文字数</div><div class=v>{chars:,}</div></div>')


def ds_warn() -> str:
    on = [x for x in ds_items() if x.get("enabled")]
    if not on:
        return '<div class=warn style="background:var(--pri-bg);color:var(--pri)">使用中のマニュアルがないため、既定の方針で校正します。</div>'
    return ""


DATASET_JS = r"""<script>
const $ = s => document.querySelector(s);
const toast = t => { const e = $('#toast'); e.textContent = t; e.classList.add('show'); clearTimeout(e._t); e._t = setTimeout(() => e.classList.remove('show'), 3000); };
const modal = $('#upModal'), pick = $('#pick'), drop = $('#drop'), flist = $('#flist'), upErr = $('#upErr');
const show = m => m.classList.add('show'), hide = m => m.classList.remove('show');
$('#q').oninput = filter;
function filter() { const v = $('#q').value.trim().toLowerCase(); document.querySelectorAll('#rows tr[data-name]').forEach(tr => tr.hidden = v && !tr.dataset.name.includes(v)); }
async function refresh() {
  let d;
  try { const r = await fetch('/api/dataset'); if (!r.ok) throw 0; d = await r.json(); } catch (_) { window._poll = setTimeout(refresh, 4000); return; }
  $('#rows').innerHTML = d.rows; $('#stats').innerHTML = d.stats; $('#warn').innerHTML = d.warn; filter();
  poll();
}
// OCR・埋め込み・ルール読み取りの処理中は数秒ごとに更新
function poll() { clearTimeout(window._poll); if (document.querySelector('#rows tr[data-task="1"]')) window._poll = setTimeout(refresh, 4000); }
poll();
const TYPE = {replace: '置換', forbid: '使わない語', regex: '正規表現', ignore: '指摘しない語', symbol_exclude: '符号でない語', symbol_prefix: '名称の前の語'};
let rulesOf = null;
document.querySelectorAll('[data-open]').forEach(b => b.onclick = () => { reset(); show(modal); });
document.querySelectorAll('[data-close]').forEach(b => b.onclick = () => hide(b.closest('.modal')));
document.querySelectorAll('.modal').forEach(m => m.addEventListener('click', e => { if (e.target === m && !m.dataset.lock) hide(m); }));
function reset() { flist.innerHTML = ''; upErr.textContent = ''; pick.value = ''; delete modal.dataset.lock; }

// 使う／使わない・編集・削除
document.addEventListener('change', async e => {
  const rl = e.target.closest('[data-rule]');
  if (rl) { await fetch(`/dataset/${rulesOf}/rule`, {method: 'POST', body: new URLSearchParams({key: rl.dataset.rule, enabled: rl.checked ? '1' : '0'})});
    toast(rl.checked ? 'このルールを使います' : 'このルールを使わないようにしました'); return; }
  const t = e.target.closest('[data-toggle]'); if (!t) return;
  await fetch(`/dataset/${t.dataset.toggle}/toggle`, {method: 'POST', body: new URLSearchParams({enabled: t.checked ? '1' : '0'})});
  toast(t.checked ? '校正に使います' : '校正に使わないようにしました'); refresh();
});
let editing = null, deleting = null;
const esc = t => t.replace(/[&<>"]/g, c => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;'}[c]));
document.addEventListener('click', async e => {
  const ch = e.target.closest('[data-chunks]');
  if (ch) { const d = await (await fetch(`/dataset/${ch.dataset.chunks}/chunks`)).json();
    $('#chName').textContent = d.name;
    $('#chList').innerHTML = d.chunks.map(c => `<div class=chunk><b>［${c.cid}］</b>${c.vec ? '' : '<span class=tag>埋め込み未計算</span> '}${esc(c.text)}</div>`).join('') || '<p class=muted>チャンクがありません</p>';
    show($('#chModal')); return; }
  const ru = e.target.closest('[data-rules]');
  if (ru) { rulesOf = ru.dataset.rules; const d = await (await fetch(`/dataset/${rulesOf}/rules`)).json();
    $('#ruName').textContent = d.name;
    $('#ruList').innerHTML = d.rules.length ? '<table class=tbl><thead><tr><th>使用</th><th>種類</th><th>内容</th><th>根拠</th><th>読み取り</th></tr></thead><tbody>' + d.rules.map((r, i) =>
      `<tr><td><label class=switch><input type=checkbox data-rule="${esc(r.key)}"${r.enabled ? ' checked' : ''}><i></i></label></td><td><span class=tag>${TYPE[r.type] || r.type}</span></td>`
      + `<td><code>${esc(r.from)}</code>${r.to ? ' → <code>' + esc(r.to) + '</code>' : ''}<div class=muted>${esc(r.note || '')}</div></td><td><span class=cite>［${esc(r.cid)}］</span></td><td class=muted>${r.by}</td></tr>`).join('') + '</tbody></table>'
      : '<p class=muted>読み取れたルールはありません。</p>';
    show($('#ruModal')); return; }
  const re = e.target.closest('[data-re]');
  if (re) { re.disabled = true; re.textContent = '処理中…';
    const r = await fetch(`/dataset/${re.dataset.re}/reprocess`, {method: 'POST'});
    toast(r.ok ? '再処理しました' : '再処理できませんでした'); refresh(); return; }
  const ed = e.target.closest('[data-edit]');
  if (ed) { const d = await (await fetch(`/dataset/${ed.dataset.edit}`)).json(); editing = ed.dataset.edit;
    $('#edName').textContent = d.name; $('#edText').value = d.text; $('#edCount').textContent = d.text.length.toLocaleString() + ' 字'; show($('#edModal')); return; }
  const del = e.target.closest('[data-del]');
  if (del) { deleting = del.dataset.del; $('#delName').textContent = del.dataset.name; show($('#delModal')); }
});
$('#edText').oninput = () => $('#edCount').textContent = $('#edText').value.length.toLocaleString() + ' 字';
$('#edSave').onclick = async () => {
  const r = await fetch(`/dataset/${editing}/text`, {method: 'POST', body: new URLSearchParams({text: $('#edText').value})});
  if (!r.ok) { toast('保存できませんでした'); return; } hide($('#edModal')); toast('保存しました'); refresh();
};
$('#delOk').onclick = async () => { await fetch(`/dataset/${deleting}/delete`, {method: 'POST'}); hide($('#delModal')); toast('削除しました'); refresh(); };

// ファイル選択・ドラッグ＆ドロップ（画面のどこに落としても受け付ける）
drop.onclick = () => pick.click();
pick.onchange = () => upload([...pick.files]);
let depth = 0;
addEventListener('dragenter', e => { if (!e.dataTransfer.types.includes('Files')) return; e.preventDefault(); depth++; $('#dropAll').classList.add('show'); drop.classList.add('over'); });
addEventListener('dragleave', e => { if (--depth <= 0) { depth = 0; $('#dropAll').classList.remove('show'); drop.classList.remove('over'); } });
addEventListener('dragover', e => e.preventDefault());
addEventListener('drop', e => { e.preventDefault(); depth = 0; $('#dropAll').classList.remove('show'); drop.classList.remove('over');
  if (e.dataTransfer.files.length) { if (!modal.classList.contains('show')) { reset(); show(modal); } upload([...e.dataTransfer.files]); } });
function sendFile(f, bar) {
  return new Promise(res => {
    const x = new XMLHttpRequest();
    x.open('POST', `/dataset/upload?name=${encodeURIComponent(f.name)}`);
    x.upload.onprogress = e => { if (e.lengthComputable) bar.style.width = (e.loaded / e.total * 100) + '%'; };
    x.onload = () => { let d = {}; try { d = JSON.parse(x.responseText); } catch (_) {} res(x.status === 200 ? {ok: true, ...d} : {ok: false, err: x.responseText}); };
    x.onerror = () => res({ok: false, err: '通信エラー'});
    x.send(f);
  });
}
async function upload(files) {
  upErr.textContent = '';
  const skipped = files.filter(f => !/\.(txt|md|docx|pdf|png|jpe?g)$/i.test(f.name) || f.name.startsWith('~$'));
  files = files.filter(f => !skipped.includes(f));
  if (!files.length) { upErr.textContent = '.txt / .md / .docx / .pdf / .png / .jpg のファイルを選んでください。'; return; }
  modal.dataset.lock = '1';
  flist.innerHTML = files.map((f, i) => `<li><div class=row><span>${f.name.replace(/</g, '&lt;')}</span><span class=muted id=s${i}>${(f.size / 1024).toFixed(0)} KB</span></div><div class=prog><i id=p${i}></i></div></li>`).join('');
  let ok = true;
  for (const [i, f] of files.entries()) {
    const r = await sendFile(f, $('#p' + i));
    $('#s' + i).textContent = r.ok ? `${r.updated ? '更新' : '追加'}（${r.chars.toLocaleString()} 字・${r.chunks} チャンク・ルール ${r.rules}）` : '失敗';
    if (!r.ok) { ok = false; upErr.textContent = `${f.name}: ${r.err}`; }
  }
  delete modal.dataset.lock;
  if (skipped.length) upErr.textContent += ` 対象外のため除外: ${skipped.map(f => f.name).join('、')}`;
  refresh();
  if (ok) { setTimeout(() => hide(modal), skipped.length ? 2500 : 900); toast('取り込みました'); }
}
</script>"""


def page_dataset() -> str:
    body = f"""<div class=head><div><h1>データセット</h1></div>
<button class=btn data-open>{ICONS["up"]}マニュアルを追加</button></div>
<div class=stats id=stats>{ds_summary()}</div><div id=warn style=margin-bottom:16px>{ds_warn()}</div>
<div class=card><div class=bar><input class=search id=q placeholder="名前で検索"></div>
<table class=tbl><thead><tr><th>名前</th><th class=hide-sm>文字数</th><th class=hide-sm>状態</th><th class=hide-sm>更新日時</th><th>使用</th><th style=text-align:right>操作</th></tr></thead>
<tbody id=rows>{ds_rows()}</tbody></table></div>
<div class=card><div class=body><details class=def><summary>既定の方針（使用中のマニュアルがないときに使います）</summary><pre>{html.escape(pr.DEFAULT_MANUAL)}</pre></details></div></div>

<div class=modal id=upModal><div class=dlg><h3>マニュアルを追加<button class=x data-close>×</button></h3><div class=in>
<div class=drop id=drop>{ICONS["up"]}<b>クリックして選択、またはドラッグ＆ドロップ</b><span class=muted>.txt / .md / .docx / .pdf / .png / .jpg（複数可）</span></div>
<input type=file id=pick multiple accept=".txt,.md,.docx,.pdf,.png,.jpg,.jpeg" hidden><ul class=flist id=flist></ul><div class=err id=upErr></div>
<p class=muted style=margin-bottom:0>同じ名前のマニュアルがあれば内容を置き換えます（使用のオン／オフはそのまま）。</p></div></div></div>

<div class=modal id=ruModal><div class="dlg wide"><h3><span>ルール: <span id=ruName></span></span><button class=x data-close>×</button></h3><div class=in><div class=chunks id=ruList></div></div></div></div>

<div class=modal id=chModal><div class="dlg wide"><h3><span>チャンク: <span id=chName></span></span><button class=x data-close>×</button></h3><div class=in><div class=chunks id=chList></div></div></div></div>

<div class=modal id=edModal><div class="dlg wide"><h3><span>編集: <span id=edName></span></span><button class=x data-close>×</button></h3><div class=in>
<textarea id=edText></textarea><div class=muted id=edCount style=text-align:right;margin-top:4px></div></div>
<div class=ft><button class="btn ghost" data-close>キャンセル</button><button class=btn id=edSave>保存</button></div></div></div>

<div class=modal id=delModal><div class=dlg><h3>マニュアルを削除<button class=x data-close>×</button></h3><div class=in>
<p style=margin-top:0>「<b id=delName></b>」を削除します。元に戻せません。</p></div>
<div class=ft><button class="btn ghost" data-close>キャンセル</button><button class="btn red" id=delOk>削除する</button></div></div></div>
<div class=dropall id=dropAll>ここにドロップしてマニュアルを追加</div>"""
    return shell("データセット", "dataset", body, DATASET_JS)


FORM_MAX = 64 * 1024            # 設定・実行のフォーム
FORM_MAX_EDIT = 5 * 1024 * 1024  # マニュアルの編集（本文を送る）


class H(BaseHTTPRequestHandler):
    def _form(self, limit: int) -> dict | None:
        """フォームを読む。大きすぎれば None（読まずに断る）"""
        n = int(self.headers.get("Content-Length", "0") or 0)
        if n > limit:
            return None
        return urllib.parse.parse_qs(self.rfile.read(n).decode("utf-8", "replace"), keep_blank_values=True)

    def _send(self, body: str, code=200, ctype="text/html; charset=utf-8"):
        b = body.encode()
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(b)))
        self.end_headers()
        self.wfile.write(b)

    def do_GET(self):
        path = urllib.parse.urlparse(self.path).path
        if path == "/":
            return self._send(page_index())
        if path.startswith("/job/"):
            if path[5:] in jobs:
                return self._send(page_job(path[5:]))
            return self._send(shell("見つかりません", "jobs", '<div class=empty>この実行記録はありません（アプリの再起動で消えます）。結果ファイルは <a href=/>明細書</a> の一覧からダウンロードできます。</div>'), 404)
        if path == "/jobs":
            return self._send(page_jobs())
        if path == "/dataset":
            return self._send(page_dataset())
        if path == "/api/dataset":
            return self._send(json.dumps({"rows": ds_rows(), "stats": ds_summary(), "warn": ds_warn()}, ensure_ascii=False), ctype="application/json")
        if path.startswith("/dataset/"):
            parts = path.split("/")[2:]
            item = next((x for x in ds_items() if x["id"] == parts[0]), None)
            if not item:
                return self._send("not found", 404)
            if parts[1:] == ["rules"]:
                rules = [dict(r, key=pr.rule_key(r)) for r in ds_rules(item["id"])]
                return self._send(json.dumps({"name": item["name"], "rules": rules}, ensure_ascii=False), ctype="application/json")
            if parts[1:] == ["chunks"]:
                chunks = [{"cid": c["cid"], "text": c["text"], "vec": bool(c.get("vec"))} for c in ds_chunks(item["id"])]
                return self._send(json.dumps({"name": item["name"], "chunks": chunks}, ensure_ascii=False), ctype="application/json")
            if parts[1:] == ["download"]:
                b = ds_text(item["id"]).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "text/plain; charset=utf-8")
                self.send_header("Content-Length", str(len(b)))
                self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{urllib.parse.quote(Path(item['name']).stem + '.txt')}")
                self.end_headers()
                return self.wfile.write(b)
            return self._send(json.dumps({"name": item["name"], "text": ds_text(item["id"])}, ensure_ascii=False), ctype="application/json")
        if path == "/api/rows":
            return self._send(json.dumps({"rows": doc_rows(), "stats": doc_stats(), "busy": any_busy()}, ensure_ascii=False), ctype="application/json")
        if path == "/settings":
            return self._send(page_settings())
        if path == "/download":
            return self._download(urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query).get("path", [""])[0])
        if path == "/healthz":
            return self._send("ok")
        self._send("not found", 404)

    def do_POST(self):
        url = urllib.parse.urlparse(self.path)
        if url.path == "/upload":
            return self._upload(urllib.parse.parse_qs(url.query))
        if url.path.startswith("/cancel/"):
            j = jobs.get(url.path[8:])
            if not j:
                return self._send("not found", 404)
            j["cancel"] = True
            return self._send("ok")
        if url.path == "/dataset/upload":
            return self._ds_upload(urllib.parse.parse_qs(url.query))
        if url.path.startswith("/dataset/"):
            parts = url.path.split("/")[2:]
            form = self._form(FORM_MAX_EDIT)
            if form is None:
                return self._send("too large", 413)
            if not any(x["id"] == parts[0] for x in ds_items()):
                return self._send("not found", 404)
            if parts[1:] == ["text"]:
                ds_update(parts[0], text=form.get("text", [""])[0].replace("\r\n", "\n").strip())
            elif parts[1:] == ["toggle"]:
                ds_update(parts[0], enabled=form.get("enabled", ["1"])[0] == "1")
            elif parts[1:] == ["delete"]:
                ds_delete(parts[0])
            elif parts[1:] == ["rule"]:
                ds_update(parts[0], rule_key=form.get("key", [""])[0], rule_on=form.get("enabled", ["1"])[0] == "1")
            elif parts[1:] == ["reprocess"]:
                try:
                    ds_update(parts[0], reprocess=True)
                except Exception as e:
                    return self._send(str(e), 400)
            else:
                return self._send("not found", 404)
            return self._send("ok")
        if self.path == "/settings":
            form = self._form(FORM_MAX)
            if form is None:
                return self._send("too large", 413)
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
        form = self._form(FORM_MAX)
        if form is None:
            return self._send("too large", 413)
        try:
            p = safe(form["path"][0])
        except Exception:
            return self._send("bad request", 400)
        jid = f"{int(time.time()*1000):x}"
        with lock:
            jobs[jid] = {"path": str(p), "status": "待機中", "log": [], "outputs": [], "mode": form.get("mode", ["full"])[0],
                         "steps": {k: {"status": "待機", "note": ""} for k, _, _ in STEPS}}
            order.append(jid)
        q.put(jid)
        if self.headers.get("X-Requested-With") == "fetch":
            return self._send(json.dumps({"job": jid}), ctype="application/json")
        self.send_response(303)
        self.send_header("Location", f"/job/{jid}")
        self.end_headers()

    def _download(self, relpath: str):
        """校正結果（docx / csv）を返す。Documents の外と、それ以外の形式は返さない"""
        try:
            p = safe(relpath)
            if not p.is_file() or p.suffix.lower() not in (".docx", ".csv"):
                raise ValueError
        except Exception:
            return self._send("not found", 404)
        ctype = {".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                 ".csv": "text/csv; charset=utf-8"}[p.suffix.lower()]
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(p.stat().st_size))
        self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{urllib.parse.quote(p.name)}")
        self.end_headers()
        with open(p, "rb") as fp:
            while chunk := fp.read(1 << 20):
                self.wfile.write(chunk)

    def _ds_upload(self, qs: dict):
        n = int(self.headers.get("Content-Length", "0"))
        try:
            name = upload_name(qs["name"][0])
            if not name.lower().endswith(DATASET_EXTS):
                raise ValueError("対応していない形式です（.txt / .md / .docx / .pdf）")
            if n > 20 * 1024 * 1024:
                raise ValueError("ファイルが大きすぎます（20MB まで）")
            item, updated = ds_put(name, self.rfile.read(n), "ファイルから取り込み")
            if not item["chars"] and not item.get("images_unread"):
                ds_delete(item["id"])
                raise ValueError("文字を取り出せませんでした")
        except Exception as e:
            return self._send(f"{e}", 400)
        self._send(json.dumps({"id": item["id"], "chars": item["chars"], "chunks": item["chunks"], "rules": item.get("rules", 0), "updated": updated}),
                   ctype="application/json")

    def _upload(self, qs: dict):
        n = int(self.headers.get("Content-Length", "0"))
        try:
            folder, name = upload_name(qs["folder"][0]), upload_name(qs["name"][0])
            if not name.lower().endswith(UPLOAD_EXTS):
                raise ValueError("対応していない形式です（.docx / .pdf のみ）")
            if n > UPLOAD_MAX:
                raise ValueError("ファイルが大きすぎます")
            dest = safe(f"{UPLOAD_DIR}/{folder}/{name}")
        except Exception as e:
            return self._send(f"bad request: {e}", 400)
        dest.parent.mkdir(parents=True, exist_ok=True)
        tmp = dest.with_name("." + dest.name + ".part")
        with open(tmp, "wb") as fp:
            left = n
            while left > 0:
                chunk = self.rfile.read(min(left, 1 << 20))
                if not chunk:
                    break
                fp.write(chunk)
                left -= len(chunk)
        if left:
            tmp.unlink(missing_ok=True)
            return self._send("upload interrupted", 400)
        tmp.replace(dest)
        self._send("ok")

    def log_message(self, *a):
        pass


if __name__ == "__main__":
    threading.Thread(target=worker, daemon=True).start()
    print(f"listening on :{PORT}, docs root {ROOT}", flush=True)
    ThreadingHTTPServer(("0.0.0.0", PORT), H).serve_forever()
