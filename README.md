# startup-db-research

スタートアップDB（企業マスタ）の構築、技術アセットとの協業マッチングDB生成、そのHTMLレポート化までを行う Claude Code / Agent Skills 用スキル。

## 3つの能力

| 能力 | トリガー | 入力 | 出力 |
|---|---|---|---|
| **① スタートアップDB作成** | 「スタートアップDBを作成して」「○社集めて」「DBに追加して」 | 社数・領域（任意） | Excel（①サマリ/②項目定義/③スタートアップDB）+ JSONL + 中間CSV |
| **② マッチングDB作成** | 「マッチングDBを作成して」+ 技術アセットの Excel/CSV | アセット棚卸しファイル + 既存スタートアップDB | 同じExcelに④マッチングDB（キー4列 + DB転記4列 + 本編30列 + 要約13列 = 51列）と⑤アセット情報シートを追加 |
| **③ マッチング結果HTMLレポート** | 「マッチング結果をHTMLにして」「レポートにして」 | マッチングExcel（④/⑤/③入り） | 1ファイル完結の静的HTML（アセット一覧クリック絞り込み + 2段組カード + 協業アイデア強調 + 詳細折りたたみ） |

- グローバルは Coresignal API、国内は API申請不要のソース（公的リスト・RSS・ニュース記事）で収集。
- スタートアップDBは業界を限定せず包括的に収集し、製造業関連の技術分類（8分類）を追加タグとして付与。
- 出力は「事実のDB（スタートアップDB）」と「評価のDB（マッチングDB）」を分離。
- マッチングは3原則（①スペック起点 / ②a・b・c判定による協業価値の足切り / ③5観点grounding）に従い、生成後は機械検証で[品質]警告を検出→最大3回の品質改善パスで是正する。

出力Excelのシート構成: **①サマリ / ②項目定義 / ③スタートアップDB / ④マッチングDB / ⑤アセット情報**（④⑤はマッチング実行時のみ）。

## ディレクトリ構成

```
SKILL.md                          スキル本体（3能力の手順・項目定義・事実性の鉄則・品質改善パス）
references/
  queries.md                      検索条件・Coresignal ES DSLクエリ・国内フィルタ
  sources.md                      使用ソースのリンク集（費用つき）
  enrich_prompt.md                エンリッチ/セルフチェック用サブエージェント指示テンプレート
  gold_example.md                 マッチングDBの良品見本2ペア(架空企業。ペア生成時に必ず読ませる)
scripts/
  coresignal_search.py            Coresignal es_dsl検索→collect→JSONL
  query_default.json              包括版クエリ
  query_addon_manufacturing.json  製造業アドオン + 8分類キーワード辞書
  prtimes_rss.py                  PR TIMES公式RSSのフィルタ→CSV
  read_assets.py                  技術アセット棚卸しExcel/CSVのパース→assets.json
  validate_db.py                  JSONLの機械検証+安全な自動補正（Excel生成前に必ず実行）
  build_startup_db.py             JSONL→Excel生成（①〜③、--matching指定で④も）
  update_matching_in_excel.py     既存マッチング用xlsxの④シートだけを再生成 + [品質]警告検証
  build_matching_html.py          マッチングExcel→閲覧用HTMLレポート（能力③）
  matching_html_style.css         HTMLレポートのスタイル（見た目調整用）
  matching_html_app.js            HTMLレポートのフィルタ挙動（アセット一覧クリック絞り込み）
templates/
  スタートアップDB_テンプレート.xlsx  出力フォーマット
```

## 使い方（概略）

```bash
# ① スタートアップDB作成（収集→検証→Excel）
python scripts/coresignal_search.py --count <N> --query <クエリ> --out output/coresignal_x.jsonl --yes
python scripts/validate_db.py --in <統合JSONL> --fix
python scripts/build_startup_db.py --in <DBのjsonl> --out <xlsx>

# ② マッチングDB作成（アセットをパース→ペア生成→④⑤を生成）
python scripts/read_assets.py --in <棚卸しファイル> --out output/assets.json
python scripts/build_startup_db.py --in <DBのjsonl> --matching <matching.jsonl> --assets output/assets.json --out <xlsx>
# 既存xlsxの④だけ再生成する場合（+ 書き込み前の[品質]検証）:
python scripts/update_matching_in_excel.py --in <matching.jsonl> --xlsx <xlsx> --check-only
python scripts/update_matching_in_excel.py --in <matching.jsonl> --xlsx <xlsx>

# ③ HTMLレポート（引数なしで近傍のマッチングExcelを自動検出）
python scripts/build_matching_html.py --in <xlsx> --out <html>
```

## セットアップ

1. このリポジトリを Claude のスキルディレクトリ配下に配置（例: `.claude/skills/startup-db-research/`）。
2. Coresignal の API キーを、スキルを実行するプロジェクト直下の `APIキー.txt` に保存（キーは本リポジトリにはコミットしないこと。`.gitignore` で除外済み）。
3. Python 依存: `openpyxl`（Excel生成）。

## 品質担保の仕組み（能力②）

弱いモデルで実行しても出力品質が落ちないよう、3層で担保する。

1. **生成前**: SKILL.md の「出力品質基準」（フィールド別の文字数レンジ・必須要素）と、架空企業による良品見本2ペア（`references/gold_example.md`）を生成エージェントに必ず読ませる。禁止語（「相乗効果」「技術融合」等の定型文）も指定。
2. **機械検証**: `update_matching_in_excel.py --check-only` が**内容レベルの[品質]警告11種**を検出する — a/b/c判定なし・定型文/禁止語・文字数不足・タイトルの「X × Y」機械結合・主要リスクが①②列挙でない・スコアの非整数/等差列/全行同一ランク・行間コピペ（企業名差し替えの定型量産）・英語スラッグ漏れ・要約欠落・定型プレースホルダー残り。`score_total`（9軸平均）と`score_sum_*`4種はスクリプトが自動計算し、モデルに算術をさせない。
3. **品質改善パス（最大3回）**: 初回生成もパス1と数え、「生成 → --check-only → 対象行修正 → 再検証 → 書き込み」をパス内で完結させる。警告0で終了。事実が不明な欄は仮テキストで埋めず空欄にする。

## 注意

- Coresignal 無料トライアルは登録から7日で失効。Multi-source は 1件=2クレジット。
- 商用DBのため「事実性の鉄則」（SKILL.md 冒頭）を厳守：記憶で書かず取得ソースにある事実のみ、不明は空欄、推定は明記、納品前セルフチェック。
- 日本企業は Coresignal のカバレッジが弱いため、国内は無料ソース主体。
- `output/`・`APIキー.txt`・`__pycache__` は `.gitignore` で除外。収集した企業データやマッチング結果（商用データ）はコミットしないこと。
