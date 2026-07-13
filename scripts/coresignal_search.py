# -*- coding: utf-8 -*-
"""Coresignal Multi-source Company API: es_dsl検索 → collect → JSONL保存。

使い方:
  python coresignal_search.py --count 100 --out output/coresignal_raw.jsonl
  python coresignal_search.py --count 50 --query my_query.json --preview-only

- APIキーはプロジェクト直下の「APIキー.txt」から読む(--key-file で変更可)。
  キーの中身を標準出力・ログに出さないこと。
- クレジット消費(Multi-source): 検索1回=2、collect 1社=2。
  実行前に消費見込みを表示し、--yes が無ければ確認を求める。
- 仕様: https://docs.coresignal.com/company-api/multi-source-company-api
  POST https://api.coresignal.com/cdapi/v2/company_multi_source/search/es_dsl
  GET  https://api.coresignal.com/cdapi/v2/company_multi_source/collect/{id}
  ヘッダー: apikey: <key>
"""
import argparse, json, os, sys, time
import urllib.request
import urllib.error

BASE = "https://api.coresignal.com/cdapi/v2"
DEFAULT_KEY_FILE = os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "APIキー.txt")
DEFAULT_QUERY = os.path.join(os.path.dirname(__file__), "query_default.json")


def read_key(path):
    with open(path, encoding="utf-8-sig") as f:
        key = f.read().strip()
    if not key:
        sys.exit(f"APIキーが空です: {path}")
    return key


def call(method, url, key, body=None, retries=3):
    data = json.dumps(body).encode() if body is not None else None
    req = urllib.request.Request(url, data=data, method=method)
    req.add_header("apikey", key)
    req.add_header("accept", "application/json")
    # Cloudflare が Python-urllib のUAをブロックする(Error 1010)ため、一般的なUAを名乗る
    req.add_header("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36")
    if data:
        req.add_header("Content-Type", "application/json")
    for attempt in range(retries):
        try:
            with urllib.request.urlopen(req, timeout=60) as res:
                return json.loads(res.read().decode())
        except urllib.error.HTTPError as e:
            msg = e.read().decode(errors="replace")[:300]
            if e.code == 429 and attempt < retries - 1:
                time.sleep(2 ** attempt)
                continue
            sys.exit(f"HTTP {e.code} {url}\n{msg}")
        except urllib.error.URLError as e:
            if attempt < retries - 1:
                time.sleep(2 ** attempt)
                continue
            sys.exit(f"接続エラー: {e}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--count", type=int, default=100, help="取得する社数(collect数)")
    ap.add_argument("--query", default=DEFAULT_QUERY, help="ES DSLクエリのJSONファイル")
    ap.add_argument("--out", default="output/coresignal_raw.jsonl")
    ap.add_argument("--key-file", default=DEFAULT_KEY_FILE)
    ap.add_argument("--preview-only", action="store_true", help="preview検索のみ(件数確認)")
    ap.add_argument("--yes", action="store_true", help="クレジット消費の確認をスキップ")
    ap.add_argument("--offset", type=int, default=0, help="検索結果IDの先頭Nをスキップ(前回取得分の回避)")
    ap.add_argument("--exclude-jsonl", action="append", default=[],
                    help="過去のcollect結果JSONL。含まれるidを再取得しない(複数指定可)")
    a = ap.parse_args()

    key = read_key(a.key_file)
    with open(a.query, encoding="utf-8") as f:
        query = json.load(f)

    if a.preview_only:
        res = call("POST", f"{BASE}/company_multi_source/search/es_dsl/preview", key, query)
        print(json.dumps(res, ensure_ascii=False, indent=2)[:3000])
        return

    est = 2 + a.count * 2
    print(f"消費見込み: 検索2 + collect {a.count}社×2 = 約{est}クレジット")
    if not a.yes:
        if input("実行しますか? [y/N]: ").lower() != "y":
            sys.exit("中止しました")

    ids = call("POST", f"{BASE}/company_multi_source/search/es_dsl", key, query)
    if isinstance(ids, dict):  # ページング形式の場合に備える
        ids = ids.get("ids") or ids.get("data") or []
    print(f"検索ヒット(返却ID数): {len(ids)}")
    exclude = set()
    for p in a.exclude_jsonl:
        with open(p, encoding="utf-8") as f:
            for line in f:
                if line.strip():
                    exclude.add(json.loads(line).get("id"))
    if exclude:
        before = len(ids)
        ids = [i for i in ids if i not in exclude]
        print(f"既取得ID除外: {before - len(ids)}件")
    ids = ids[a.offset: a.offset + a.count]

    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    n = 0
    with open(a.out, "w", encoding="utf-8") as f:
        for i, cid in enumerate(ids, 1):
            rec = call("GET", f"{BASE}/company_multi_source/collect/{cid}", key)
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")
            n += 1
            if i % 10 == 0:
                print(f"  {i}/{len(ids)}")
            time.sleep(0.1)  # レート制限(54req/s)に対し十分保守的
    print(f"保存: {a.out} ({n}社) / 消費クレジット概算: {2 + n * 2}")
    print("run_log.md に クエリ・日付・件数 を追記してください(references/queries.md 参照)")


if __name__ == "__main__":
    main()
