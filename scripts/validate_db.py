# -*- coding: utf-8 -*-
"""スタートアップDB/マッチングDBのJSONLを機械検証(+安全な自動補正)する。

使い方:
  python validate_db.py --in output/startups_YYYYMMDD.jsonl            # 検証のみ
  python validate_db.py --in output/startups_YYYYMMDD.jsonl --fix      # 安全な補正を適用して上書き
  python validate_db.py --matching output/matching_YYYYMMDD.jsonl      # マッチングDBの検証

目的: エンリッチ担当(LLM)の出力ブレを、Excel生成前に決定的に検出する。
Excel生成(build_startup_db.py)の前に必ず実行すること。

検証内容(スタートアップDB):
- 必須フィールドの存在(startup_id/industry_category/name/country/fit/source_urls/info_date/collect_route)
- 語彙: industry_category(20分類)/mfg_tech_category(8分類orNull)/stage_normalized(7語彙orNull)/fit(○/△)
- 型: funding_usd・founded_yearは数値orNull、リスト型フィールドはlist
- 重複: name(小文字)/name_ja/startup_id
--fix で自動補正するもの(安全なもののみ):
- mfg_tech_categoryの複数指定(";"区切り/リスト)→先頭の有効値
- industry_categoryの部分一致(例:「モビリティ」→「モビリティ・自動車」)
- 数値フィールドに入った文字列数値→数値化(不能ならNoneにして報告)
※fitの不正値・重複・必須欠落は自動補正しない(人/上位モデルの判断が必要なため報告のみ)

検証内容(マッチングDB): score_*が0-10の数値orNull、必須キー(startup_id/asset_id/coop_type/title)の存在。
終了コード: 0=問題なし(または全て補正済み) / 1=要対応の問題あり
"""
import argparse, json, sys

CATS = ["AI・データ基盤", "エンタープライズSaaS・業務DX", "フィンテック・保険", "ヘルスケア・医療機器",
        "バイオ・創薬", "モビリティ・自動車", "宇宙・航空・防衛", "エネルギー・クリーンテック",
        "素材・化学", "半導体・エレクトロニクス", "ロボティクス・ドローン", "製造DX・産業機器",
        "物流・サプライチェーン", "農業・食品", "建設・不動産・インフラ", "小売・EC・消費者サービス",
        "HR・教育", "セキュリティ・通信", "エンタメ・メディア・ゲーム", "その他"]
MFG = ["AI・データ(シミュレーション/CAE)", "センシング", "検査・品質", "予知保全",
       "ロボティクス・FA", "エッジAI", "材料・MI", "熱・パワエレ"]
STAGES = ["シード", "シリーズA", "シリーズB", "シリーズC", "シリーズD以降", "グラント・デットのみ", "不明"]
REQUIRED = ["startup_id", "industry_category", "name", "country", "fit", "source_urls", "info_date", "collect_route"]
LIST_FIELDS = ["key_investors", "target_industries", "notable_customers_partners", "tech_keywords", "source_urls"]


def load_jsonl(path):
    rows = []
    with open(path, encoding="utf-8") as f:
        for i, line in enumerate(f, 1):
            if line.strip():
                try:
                    rows.append(json.loads(line))
                except json.JSONDecodeError as e:
                    print(f"[致命] {i}行目: JSONとして読めない: {e}")
                    sys.exit(1)
    return rows


def to_num(v):
    if v is None or isinstance(v, (int, float)):
        return v, False
    try:
        return float(str(v).replace(",", "")), True
    except ValueError:
        return None, True


def validate_db(path, fix):
    rows = load_jsonl(path)
    errors, fixes = [], []
    seen = {}
    for idx, d in enumerate(rows, 1):
        label = d.get("name_ja") or d.get("name") or f"{idx}行目"
        for k in REQUIRED:
            if not d.get(k):
                errors.append(f"必須欠落: {label} / {k}")
        ic = d.get("industry_category")
        if ic not in CATS:
            near = next((c for c in CATS if ic and (str(ic) in c or c in str(ic))), None)
            if fix and near:
                d["industry_category"] = near
                fixes.append(f"業界分類補正: {label} 「{ic}」→「{near}」")
            else:
                errors.append(f"業界分類が語彙外: {label} 「{ic}」" + (f"(候補:{near})" if near else ""))
        mc = d.get("mfg_tech_category")
        if mc not in MFG + [None]:
            cand = None
            if isinstance(mc, list) and mc:
                cand = str(mc[0])
            elif isinstance(mc, str):
                cand = mc.split(";")[0].split("；")[0].strip()
            near = next((c for c in MFG if cand and (cand in c or c in cand)), None)
            if fix:
                d["mfg_tech_category"] = near
                fixes.append(f"製造分類補正: {label} {mc!r}→{near!r}")
            else:
                errors.append(f"製造分類が語彙外: {label} {mc!r}")
        sn = d.get("stage_normalized")
        if sn not in STAGES + [None]:
            errors.append(f"ステージ統一語彙外: {label} 「{sn}」(有効: {'/'.join(STAGES)})")
        if d.get("fit") not in ("○", "△"):
            errors.append(f"fitが不正: {label} {d.get('fit')!r}(○/△のみ。理由はfit_noteへ)")
        for k in ("funding_usd", "founded_year"):
            v, changed = to_num(d.get(k))
            if changed:
                if fix:
                    d[k] = v
                    fixes.append(f"{k}数値化: {label} → {v}")
                else:
                    errors.append(f"{k}が数値でない: {label} {d.get(k)!r}")
        for k in LIST_FIELDS:
            v = d.get(k)
            if v is not None and not isinstance(v, list):
                if fix and isinstance(v, str):
                    d[k] = [v]
                    fixes.append(f"{k}をリスト化: {label}")
                else:
                    errors.append(f"{k}がリストでない: {label}")
        for key in [("name", (d.get("name") or "").lower()), ("startup_id", d.get("startup_id"))]:
            if key[1]:
                if key[1] in seen:
                    errors.append(f"重複({key[0]}): {label} と {seen[key[1]]}")
                else:
                    seen[key[1]] = label
    if fix and fixes:
        with open(path, "w", encoding="utf-8") as f:
            for d in rows:
                f.write(json.dumps(d, ensure_ascii=False) + "\n")
    return rows, errors, fixes


def validate_matching(path):
    rows = load_jsonl(path)
    errors = []
    for idx, d in enumerate(rows, 1):
        label = d.get("startup_name") or f"{idx}行目"
        for k in ("startup_id", "asset_id", "coop_type", "title"):
            if not d.get(k):
                errors.append(f"必須欠落: {label} / {k}")
        for k, v in d.items():
            if k.startswith("score_") and v is not None:
                if not isinstance(v, (int, float)):
                    errors.append(f"スコアが数値でない: {label} / {k}={v!r}(「8/10」等の文字列は禁止)")
                elif not (0 <= v <= 10):
                    errors.append(f"スコアが0-10範囲外: {label} / {k}={v}")
    return rows, errors


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="inp", help="スタートアップDBのJSONL")
    ap.add_argument("--matching", help="マッチングDBのJSONL")
    ap.add_argument("--fix", action="store_true", help="安全な自動補正を適用して上書き")
    a = ap.parse_args()
    if not a.inp and not a.matching:
        sys.exit("--in または --matching を指定してください")
    ng = 0
    if a.inp:
        rows, errors, fixes = validate_db(a.inp, a.fix)
        print(f"[スタートアップDB] {len(rows)}件検証")
        for x in fixes:
            print("  補正:", x)
        for x in errors:
            print("  要対応:", x)
        print(f"  → 補正{len(fixes)}件 / 要対応{len(errors)}件")
        ng += len(errors)
    if a.matching:
        rows, errors = validate_matching(a.matching)
        print(f"[マッチングDB] {len(rows)}件検証")
        for x in errors:
            print("  要対応:", x)
        print(f"  → 要対応{len(errors)}件")
        ng += len(errors)
    sys.exit(1 if ng else 0)


if __name__ == "__main__":
    main()
