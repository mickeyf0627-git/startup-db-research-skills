# 検索条件・クエリ定義

方針: **収集は包括的(業界を絞らない)**。スタートアップ条件のみで検索し、業界分類はLLMで後付けする。
そのうえで、利用目的(製造業の技術マッチング)から**製造業関連を厚めに取りたい場合**は、追加クエリ(アドオン)を組み合わせる。

## 1. スタートアップ条件のクエリ翻訳

| 定義 | Coresignal(グローバル) | 国内(API不使用) |
|---|---|---|
| 未上場 | `must_not: is_public=true` | 上場・買収をニュース/サイトで確認 |
| 設立10年以内(ディープテック15年) | `founded_year >= 2016`(ディープテック厚めなら`>= 2011`に緩和) | 公式サイト・登記情報 |
| 外部エクイティ調達あり | `last_funding_round.announced_date >= 2024-01-01`(鮮度重視。網羅重視なら2021-01-01に緩和) | 調達プレスリリースの存在(=収集起点なので自動的に満たす) |
| 従業員500名未満 | `employees_count <= 500` | サイト・採用ページの概算 |

## 2. Coresignal ES DSL クエリ

- エンドポイント: `POST https://api.coresignal.com/cdapi/v2/company_multi_source/search/es_dsl`
- ヘッダー: `apikey: <APIキー.txtの中身>` / `Content-Type: application/json`
- 消費: 検索1回=2クレジット、collect 1社=2クレジット(Multi-source)

### 既定クエリ(包括版): `scripts/query_default.json`

業界・キーワードの絞り込みなし。スタートアップ条件のみ+直近調達日の降順ソート(=「最近調達したスタートアップ」から順に取れる)。
※sort句が拒否された場合(400エラー)はsortを外して実行し、その旨をrun_logに記録する。

### 製造業アドオン: `scripts/query_addon_manufacturing.json`

製造業関連を厚めにしたいとき、既定クエリの `must` 配列にこのファイルの `bool.should` ブロックを追加する(minimum_should_match=1)。中身:

- `industry` が製造関連(Industrial Automation / Machinery Manufacturing / Semiconductors / Robotics Engineering / Motor Vehicle (Parts) Manufacturing 等)
- または `description` / `categories_and_keywords` に製造系技術キーワード(robotics, machine vision, predictive maintenance, surrogate model, CAE, digital twin, edge AI, materials informatics, thermal management, power electronics, in-cabin sensing, driver monitoring, quality inspection 等)

**推奨の配分**: 包括的に集める場合も、目標N社のうち3〜5割は製造業アドオン付きクエリで取得し、残りを包括版で取得する(利用目的とのバランス)。配分はユーザー指定があれば従う。

### 業界分類別の検索キーワード(領域指定があった場合のshould句差し替え用)

| 業界分類 | 英語キーワード例 |
|---|---|
| AI・データ基盤 | foundation model, LLM, data platform, MLOps, AI agent |
| エンタープライズSaaS・業務DX | enterprise software, workflow automation, ERP, legal tech |
| フィンテック・保険 | fintech, payments, lending, insurtech |
| ヘルスケア・医療機器 | digital health, medical device, diagnostics |
| バイオ・創薬 | biotech, drug discovery, synthetic biology |
| モビリティ・自動車 | mobility, autonomous driving, EV, ADAS, in-cabin sensing, driver monitoring |
| 宇宙・航空・防衛 | space, satellite, aerospace, defense |
| エネルギー・クリーンテック | clean energy, battery, hydrogen, carbon capture |
| 素材・化学 | advanced materials, materials informatics, specialty chemicals |
| 半導体・エレクトロニクス | semiconductor, chip design, photonics, edge AI accelerator |
| ロボティクス・ドローン | robotics, robot foundation model, bin picking, AMR, drone |
| 製造DX・産業機器 | industrial automation, machine vision, quality inspection, predictive maintenance, digital twin, CAE, surrogate model, additive manufacturing |
| 物流・サプライチェーン | logistics, supply chain, warehouse automation |
| 農業・食品 | agtech, foodtech, alternative protein |
| 建設・不動産・インフラ | construction tech, proptech, infrastructure |
| 小売・EC・消費者サービス | e-commerce, retail tech, D2C |
| HR・教育 | HR tech, edtech, talent |
| セキュリティ・通信 | cybersecurity, network, telecom |
| エンタメ・メディア・ゲーム | gaming, media, creator economy |
| (製造業関連8分類の詳細キーワードは query_addon_manufacturing.json を参照) | |

### 実行フロー(クレジット節約)

1. まず `search/es_dsl/preview` で件数と上位のプレビューを確認(低コスト)
2. 件数が目標Nに対して過大なら条件を絞る(設立年・調達日を厳しく)、過小なら緩める(2011年/2021年へ)
3. 本検索 → 上位N件のIDを `GET /v2/company_multi_source/collect/{id}` で取得

## 3. 国内: RSS・記事のフィルタ条件

### PR TIMES RSS (`https://prtimes.jp/index.rdf`)

タイトルに対するキーワードフィルタ(scripts/prtimes_rss.py の既定):

- 含む: `資金調達 / 億円を調達 / 万円を調達 / シリーズA / シリーズB / シリーズC / プレシリーズ / シードラウンド / 第三者割当増資 / 資金調達を実施 / エクイティ`
- 除外(LLM二次判定): 上場企業の公募増資、投資ファンド組成のみのリリース

### KEPPLE / BRIDGE 週次まとめ

- KEPPLE「資金調達Weekly」・BRIDGE「国内スタートアップ資金調達振り返り」を直近4〜8週分読む
- 抽出項目: 社名 / 調達額 / ラウンド / 投資家 / 事業一行説明 / 記事URL
- **本文は保存せず、事実データのみ抽出**(規約対応)
- 業界は絞らず全案件を抽出→LLMで20分類に振り分け(製造業関連は8分類の追加タグも付与)

### 中間CSVの出力(必須)

抽出した国内候補は必ず `output/domestic_candidates_YYYYMMDD.csv` に保存する。
列: `社名, 調達額, ラウンド, 投資家, 事業一行, 記事URL, 取得経路(prtimes-rss/kepple/bridge/j-startup/sbir-nedo)`

## 4. 実行記録ルール(再現性)

- 実行したクエリJSON・実行日・ヒット件数・取得件数・消費クレジットを `output/run_log.md` に追記する。
- クエリを変更した場合は該当JSONファイルを更新し、変更理由をrun_logに記す。
