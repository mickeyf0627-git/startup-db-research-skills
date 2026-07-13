# エンリッチ用サブエージェント指示テンプレート

サブエージェント(Agentツール)でエンリッチを行うときは、**以下のテンプレートをそのままコピーし、【】のプレースホルダだけ置換して使う**こと。自分で指示文を作文しない(モデル・実行者によるブレを防ぐため)。

- 1エージェントの担当は**最大8社**。9社以上は複数エージェントに分割する。
- 出力ファイル名は `output/enrich_<バッチ名>.json` で重複させない。

---

## テンプレートA: グローバル企業(Coresignal由来)

```
あなたは商用スタートアップDBのエンリッチ担当。今日は【YYYY-MM-DD】。作業ディレクトリ: 【プロジェクトの絶対パス】

【必読】まず `.claude/skills/startup-db-research/SKILL.md` の「事実性の鉄則」と「スタートアップDBの項目(40項目)」を読み厳守すること。

【担当】`【slimファイルのパス】` から次の【N】社: 【社名をカンマ区切りで列挙】

【手順】各社: slimレコードの事実をベース→公式サイトをWebFetch→必要ならWebSearch1回まで→40項目のJSONを作成。
- 業界分類20/製造技術分類8はSKILL.mdの定義の語彙をそのまま使う(製造技術分類は必ず1つだけ。複数該当時は最も中核的な1つ)
- stage_normalized: シード/シリーズA/B/C/D以降/グラント・デットのみ/不明 のいずれか
- funding_usd: 百万USDの数値のみ(換算: EUR=1.1, GBP=1.27, JPY=1/150)。累計不明ならnull
- market_size/CAGR系: 検索1回で出典が見つかった場合のみ記入(見つからなければnull。無理をしない)
- fit判定は定義6基準のみ(○か△の2値。理由はfit_noteに)。上場・買収・事業停止を確認
- 企業名はタグライン・記号を除いた正式名に整形
- source_urls: "coresignal:multi-source"+使用した全URL
- info_date="【YYYY-MM-DD】", collect_route="coresignal"
【出力】【N】社のJSON配列を `【出力ファイルパス】` にWrite。最終メッセージは「社数/fit△と理由/製造技術分類を付けた社数」のみ。
```

## テンプレートB: 国内企業(ニュース・公的リスト由来)

```
あなたは商用スタートアップDBのエンリッチ担当。今日は【YYYY-MM-DD】。作業ディレクトリ: 【プロジェクトの絶対パス】

【必読】まず `.claude/skills/startup-db-research/SKILL.md` の「事実性の鉄則」と「スタートアップDBの項目(40項目)」を読み厳守すること。

【担当】`【候補ファイルのパス】` から次の【N】社: 【社名をカンマ区切りで列挙】

【手順】各社: 候補レコードの事実(調達額・投資家・事業一行・記事URL)をベース→WebSearchで公式サイトを見つけWebFetch→40項目のJSONを作成。
- 国内共通: name_ja=正式社名, name=英語名/ローマ字, country="日本", startup_id="jp-"+ローマ字/ドメイン名, collect_route=候補レコードの値
- 調達情報が候補レコードでnullの会社は、公式サイト/検索で直近調達を1件確認できたら記載(できなければnull)
- 業界分類20/製造技術分類8はSKILL.mdの定義の語彙をそのまま使う(製造技術分類は必ず1つだけ)
- stage_normalized: シード/シリーズA/B/C/D以降/グラント・デットのみ/不明 のいずれか
- funding_usd: 百万USDの数値のみ(JPY=1/150換算)
- market_size/CAGR系: 検索1回で出典が見つかった場合のみ(無ければnull)
- fit判定は定義6基準のみ(○/△)。上場・買収・大企業子会社を確認
- info_date="【YYYY-MM-DD】"
【出力】【N】社のJSON配列を `【出力ファイルパス】` にWrite。最終メッセージは「社数/fit△と理由/製造技術分類を付けた社数」のみ。
```

## テンプレートC: 納品前セルフチェック

```
あなたは商用スタートアップDBの品質検査担当(作成には非関与)。今日は【YYYY-MM-DD】。作業ディレクトリ: 【プロジェクトの絶対パス】

`【最終JSONLのパス】` から次の【3〜5】社のレコードを読み、各フィールドの値がsource_urlsのソースで確認できるかを突合検査する: 【社名を列挙(新規追加分から無作為)】

検査方法:
1. 各社のsource_urlsのURL("coresignal:multi-source"は除く)をWebFetchで開く
2. 主要フィールド(設立年/所在地/調達額・ラウンド・時期/投資家/プロダクト名/事業内容/顧客名)がソース記載と一致するか確認
3. 「(推定)」付きは矛盾がなければ許容。coresignalのみが出典のフィールドは「API由来・未確認」に分類(不一致ではない)
4. 不一致は「フィールド名/レコードの値/ソースの記載」を具体的に列挙

出力(最終メッセージ): 各社の「一致/不一致/API由来」内訳と不一致の詳細、最後に総合判定(合格/要修正)1行。ファイル書き込み不要。
```

## テンプレートD: マッチング(能力2・アセット×スタートアップ)

SKILL.md「マッチング品質の3原則」を必ず適用する。1エージェント=1アセット担当が扱いやすい。

```
あなた自身がWebで調べて書く技術アセット×スタートアップの協業マッチング担当。**他のエージェントは起動しない。** 今日は【YYYY-MM-DD】。作業ディレクトリ: 【プロジェクトの絶対パス】

【最初に】ToolSearchでWebSearch,WebFetchを読み込む。`.claude/skills/startup-db-research/SKILL.md`の「事実性の鉄則」「マッチング品質の3原則」「マッチングDBの項目」を読む。
【入力】`【assets.jsonのパス】`の【何番目】のアセット「【アセット名】」を熟読(機能ツリー・スペック・WANT・NG条件)。`【スタートアップDBのjsonl】`(【N】社)。

【任務】このアセットと協業可能性のある**上位【5〜10】社**を選び、協業アイデアを1件ずつ生成する。
- 原則1: **アセットの最も尖った強み(【強みを1〜2点明記】)を起点**に「その強みが効く相手」を探す(タグ一致で選ばない)
- 原則2: 各ペアは (a)性能突破 / (b)コスト等1/10 / (c)別セグメント転用 のどれを満たすか判定。満たさない小幅改善はCランク(総合6点未満)に抑える
- 原則3: agent_rationaleに5観点(主要市場/競合/顧客課題/トレンド/規制)の最低2つを事実で言及。市場データは1ペアWebSearch1回まで・出典確認時のみ
- **正直採点**: 良い相手がいなければ低スコア可。無理に高得点にしない

【各レコードのJSONフィールド(この名前で。スコアは全て0-10の数値)】
startup_id, startup_name, asset_id, coop_type, title, tags(;区切り), summary, market_size_list, cagr_list, target_customer, pain, solution, revenue_model, idea_strength, core_tech_used("アセット側: …/スタートアップ側: …"), score_differentiation, score_tech_feasibility, score_asset_fit, score_industry_advantage, score_industry_pain_fit, unused_reason_barriers, core_improvement, score_novelty, score_mission_fit, market_size_detail, score_market_size, cagr_detail, score_cagr, score_total(採点軸の平均・小数1桁), agent_rationale(**満たした判定式a/b/cを明記+5観点2つ以上**), key_risks, および要約ブロック sum_concept/sum_market/sum_customer/sum_pain/sum_solution/sum_market_size/sum_cagr/sum_uniqueness/sum_mission と score_sum_uniqueness(=score_novelty)/score_sum_mission(=score_mission_fit)/score_sum_market_size(=score_market_size)/score_sum_growth(=score_cagr)
※スタートアップの事業概要等の転記列はbuild_startup_db.pyが自動結合するので書かない
【事実性】相手企業の記述はDBレコードの記載範囲のみ。捏造禁止。

【出力】JSON配列を `【出力パス】` にWrite。最終メッセージは「選定社数とscore_total、各ペアが満たした判定式(a/b/c)」のみ。
```

## 運用メモ

- エージェントが「セッション上限」等のエラーで終了しても、**出力ファイルは書き込み済みのことが多い**。再実行前に必ず出力ファイルの存在と件数を確認する。
- 完了したら必ず `python scripts/validate_db.py --in <統合JSONL> --fix` を実行し、「要対応」が0になるまで修正してからExcelを生成する。
