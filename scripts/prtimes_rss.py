# -*- coding: utf-8 -*-
"""PR TIMES公式RSSから資金調達リリース候補を抽出してCSV保存。

使い方:
  python prtimes_rss.py --out output/prtimes_candidates.csv

- 公式RSS: https://prtimes.jp/index.rdf (直近リリースのみ流れるため、バックフィルには
  KEPPLE Weekly / BRIDGE週次まとめの記事をWebFetchで読むこと)
- 規約対応: 本文は取得・保存しない。タイトル/URL/日付のみ扱う。
- 出力CSV列: date, title, url, matched_keyword
"""
import argparse, csv, os, re, sys
import urllib.request
import xml.etree.ElementTree as ET

RSS_URL = "https://prtimes.jp/index.rdf"
KEYWORDS = [
    "資金調達", "億円を調達", "万円を調達", "シリーズA", "シリーズB", "シリーズC",
    "プレシリーズ", "シードラウンド", "第三者割当増資", "資金調達を実施", "エクイティ",
]
NS = {
    "rss": "http://purl.org/rss/1.0/",
    "dc": "http://purl.org/dc/elements/1.1/",
}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="output/prtimes_candidates.csv")
    ap.add_argument("--rss-url", default=RSS_URL)
    a = ap.parse_args()

    req = urllib.request.Request(a.rss_url, headers={"User-Agent": "Mozilla/5.0 (research; low-frequency)"})
    with urllib.request.urlopen(req, timeout=60) as res:
        xml = res.read()
    root = ET.fromstring(xml)

    items = root.findall("rss:item", NS) or root.findall(".//item")
    rows = []
    for it in items:
        title = (it.findtext("rss:title", "", NS) or it.findtext("title", "") or "").strip()
        link = (it.findtext("rss:link", "", NS) or it.findtext("link", "") or "").strip()
        date = (it.findtext("dc:date", "", NS) or it.findtext("pubDate", "") or "").strip()
        hit = next((k for k in KEYWORDS if k in title), None)
        if hit:
            rows.append({"date": date, "title": title, "url": link, "matched_keyword": hit})

    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    with open(a.out, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["date", "title", "url", "matched_keyword"])
        w.writeheader()
        w.writerows(rows)
    print(f"RSS {len(items)}件中、資金調達候補 {len(rows)}件 → {a.out}")
    if not rows:
        print("ヒット0件。RSSは直近数時間分のみのため、時間を置いて再実行するか、KEPPLE/BRIDGEの週次まとめでバックフィルしてください")


if __name__ == "__main__":
    main()
