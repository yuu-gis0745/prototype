# プロジェクト名

森林調査データ入力・検査効率化ツール

---

## 📌 概要（30秒で分かるこのツール）

森林調査票（xlsx／PDF）の入力内容を自動チェックし、エラーログ出力・GIS用ファイル変換・Slack通知までを一括で行う業務効率化ツールです。測量・森林調査の実務経験（15年超）を基に、Python・生成AI・Difyの活用力を示すために制作したプロトタイプです。

[![Python](https://img.shields.io/badge/Python-3776AB?style=flat&logo=python&logoColor=white)](https://www.python.org/)
[![FastAPI](https://img.shields.io/badge/FastAPI-009688?style=flat&logo=fastapi&logoColor=white)](https://fastapi.tiangolo.com/)
[![Dify](https://img.shields.io/badge/Dify-0033FF?style=flat&logo=dify&logoColor=white)](https://dify.ai/)
[![OpenAI API](https://img.shields.io/badge/OpenAI_API-412991?style=flat)](https://platform.openai.com/)
[![Google Cloud Vision](https://img.shields.io/badge/Google%20Cloud%20Vision-4285F4?style=flat&logo=googlecloud&logoColor=white)](https://cloud.google.com/vision)
[![QGIS](https://img.shields.io/badge/QGIS-589632?style=flat&logo=qgis&logoColor=white)](https://qgis.org/)

**主な成果**
- 既存ルール＋LLMによる入力エラーの自動検出（追加候補抽出含む）
- PDFのOCR読み取り→Excel形式への復元
- GIS用ファイルへの自動変換、Slack通知までの一連の自動化

👉 詳細な背景・課題・技術選定理由は以下に続きます。

---

## 目次

- [1．プロジェクト概要](#1プロジェクト概要)
- [2．制作背景](#2制作背景)
- [3．解決したい課題](#3解決したい課題)
- [4．主な機能](#4主な機能)
- [5．使用技術・目的](#5使用技術目的)
- [6．フォルダ構成](#6フォルダ構成)
- [7．処理内容](#7処理内容)
  - [7-1．処理内容の概要](#7-1処理内容の概要)
  - [7-2．各処理段階の概要](#7-2各処理段階の概要)
  - [7-3．処理方法](#7-3処理方法)
- [8．実行方法](#8実行方法)
  - [8-1．ローカルでの実行方法](#8-1ローカルでの実行方法)
  - [8-2．Renderにおけるデプロイ設定](#8-2renderにおけるデプロイ設定)
  - [8-3．Difyとの連携](#8-3difyとの連携)
- [9．処理結果・画面例](#9処理結果画面例)
  - [9-1．想定される処理フローの変化](#9-1想定される処理フローの変化)
  - [9-2．主要な処理結果](#9-2主要な処理結果)
  - [9-3．Slack通知](#9-3slack通知)
- [10．工夫した点](#10工夫した点)
- [11．生成AIに支援してもらった部分・自分で担当した部分](#11生成aiに支援してもらった部分自分で担当した部分)
  - [11-1．生成AIに支援してもらった部分](#11-1生成aiに支援してもらった部分)
  - [11-2．自分で担当・確認した部分](#11-2自分で担当確認した部分)
- [12．現時点における成果](#12現時点における成果)
- [13．現時点における課題](#13現時点における課題)
- [14．今後の改善点・取り入れたい技術](#14今後の改善点取り入れたい技術)
- [15．参考資料・使用データ](#15参考資料使用データ)

---

## 1．プロジェクト概要

このツールは、森林調査で得られた現地調査結果の手入力による整理と入力結果の検査、GIS用ファイルの作成の効率化を目的として制作したものです。

---

## 2．制作背景

測量業務では、調査後の測量データを指定のExcel様式へ手入力する作業があり、転記ミスや誤判読の発生、入力後の検査に時間を要する点を課題として感じていました。
そこで、私が業務で経験した森林調査をテーマとして、ITスクールで学習した生成AI・Python・Difyを活用し、架空の森林調査票を用いた測量調査データの入力支援と入力後の検査、GIS用ファイルへの変換・作成を行うツールを、業務改善の検証を目的としたプロトタイプとして制作いたしました。

---

## 3．解決したい課題

- 調査結果のExcel様式への入力作業の時間短縮、入力漏れ・文字や数値の誤入力・転記ミスの減少
  * 調査結果は、現地において紙の調査票へ記載したものが中心です。
  * 紙資料は、記帳者の文字のくせ字・乱雑さ、資料そのものの汚れや折れ曲がりがあり文字や数値の誤入力が発生しやすいです。

- 入力結果の検査、エラー修正の精度向上と時間短縮
  * Excel様式で入力した調査結果を、PC画面上もしくは印刷して確認しております。
  * PC画面上や印刷物による目視や読み合わせによる確認は、時間を要しかつチェック漏れ発生の可能性がございます。

- GIS用ファイルの作成の効率化
  * 座標値から、GISソフト上に展開して位置確認を行っております。
  * Excel様式から、GIS用ファイルへ変換するのに別途編集作業が必要であり、煩雑でかつ時間を要しております。

---

## 4．主な機能

このツールで実装した主な機能は、以下の通りです。

| 機能 | 内容 |
|---|---|
| 入力結果の取り込み | xlsx・PDF形式の森林調査票(調査結果を入力済み)から、入力結果を取り込む。 |
| OCR | PDFの場合には、AI OCR機能を用いて処理を実行する。 |
| PDF → Excel形式への復元 | AI OCR機能を用いて抽出したテキストを特定のExcel形式に復元する。 |
| 入力内容のチェック | 既存エラールール(以下既存エラー)が記載されたcsvファイルと、森林調査票の調査項目のセル位置を記載したxlsxから入力内容を検査する。 |
| エラーログの出力 | 入力内容の検査の結果、エラーがあった場合にはcsv形式のエラーログを出力する。 |
| 追加エラー候補のチェック | 既存エラーにない新規のエラーの可能性を検査し、追加エラー候補として抽出する。 |
| GIS用ファイルへの変換・出力 | 入力結果の検査の結果、既存エラーがない場合には調査地点ごとで集計を行い、GISソフト上で表示可能なファイルのフォーマット形式へ変換・出力を行う。 |
| Slackへの通知 | LLMに、エラーの有無と、既存エラー内容と修正指示、追加エラー候補の確認内容についての文章を作成させ、Slackへ投稿する。 |

---

## 5．使用技術・目的

主な機能を実装するために使用した技術と、その技術を使用した主な目的は以下の通りです。各技術がどの処理段階で使われているかは、[7-2．各処理段階の概要](#7-2各処理段階の概要)でご確認いただけます。

| 技術 | 主な目的 |
|---|---|
| Python 3.12以上 | xlsxの読み取り・入力内容の検査・エラーログ出力・GIS用ファイル変換など、複雑な処理ロジックを実装 |
| FastAPI | DifyからPythonコードを呼び出すためのAPIとして公開 |
| Dify(Webクラウド版) | ファイルのアップロードから出力までの一連の処理をワークフローとして統合 |
| ChatGPT(OpenAI API) | コード開発・改善案作成の支援、追加エラー候補の抽出、Slack投稿文の作成、README構成の作成支援 |
| Claude Code | コード開発・テスト実施・エラー原因の特定支援、READMEのレビュー・改善案作成 |
| openpyxl | 森林調査票(xlsx)のセル位置取得、読み書き |
| Google Cloud Vision API | PDFからの文字・数字のOCR抽出 |
| PyMuPDF | OCR結果に対する追加の文字・数字抽出(表形式対応) |
| QGIS 3.44.0 | 出力したGIS用ファイルの表示確認 |
| Render | FastAPIの外部公開(Dify連携用) |
| Slack | エラー内容・修正指示・追加エラー候補確認の通知 |
| GitHub / Visual Studio Code | コード開発・バージョン管理・README公開 |

---

## 6．フォルダ構成

主要フォルダは以下の通りです：`api`(FastAPI)／`dify`(Difyワークフロー)／`gis`(QGIS確認用)／`master`(検査ルール・テンプレート)／`samples`(入出力サンプル)。詳細なファイル構成は以下を展開してください。

<details>
<summary>フォルダ構成の詳細を見る</summary>

```
forest_survey_check_tool/
├── README.md  # プロジェクト説明
├── requirements.txt  # ライブラリ一覧
├── .python-version  # Pythonバージョン固定用ファイル
├── .gitignore  # GitHubに含めないファイル設定
│
├── api/
│   └── main.py  # FastAPIのメイン処理
│
├── dify/
│   └── forest_survey_check_tool.yml  # Difyのymlファイル
│
├── docs/
│   ├── deploy_render.md  # Renderを用いたデプロイの方法をまとめた資料
│   ├── screenshots.md  # 処理結果・画面例の補足資料
│   └── images/  # README用画像
│
├── gis/
│   ├── forest_survey_check.qgz  # QGIS表示確認用プロジェクトファイル
│   └── data/
│       ├── gis_plot_summary.csv  # QGIS表示用CSV
│       └── qgis_plot_summary.gpkg  # QGIS表示用GeoPackage(QGISで確認しやすいように、GeoPackage形式のサンプルデータも格納しております。)
│
├── master/
│   ├── template_forest_survey.xlsx  # 森林調査票の原本
│   ├── check_rules_forest_survey.csv  # 既存のエラールールを定義したファイル
│   └── forest_survey_cell_mapping.xlsx  # 森林調査票の調査項目のセル位置対応表
│
└── samples/
    ├── input/
    │   ├── sample_forest_survey_3plots_with_error.pdf  # 入力用サンプルPDF(入力ミスあり)
    │   ├── sample_forest_survey_3plots_no_error.pdf # 入力用サンプルPDF(入力ミスなし)
    │   ├── sample_forest_survey_3plots_with_error.xlsx  # 入力用サンプルExcel(入力ミスあり)
    │   └── sample_forest_survey_3plots_no_error.xlsx  # 入力用サンプルExcel(入力ミスなし)
    │
    └── output/
        ├── error_log_from_pdf_with_error.csv  # エラーログ(入力ファイルがPDF)
        ├── error_log_from_pdf_no_error.csv  # エラーログ(入力ファイルがPDF)
        ├── error_log_from_xlsx_with_error.csv  # エラーログ(入力ファイルがxlsx)
        ├── error_log_from_xlsx_no_error.csv  # エラーログ(入力ファイルがxlsx)
        ├── gis_plot_summary.csv  # GIS表示用ファイル
        ├── ocr_result_with_error.csv  # PDFの文字・数字抽出結果(元の森林調査票にエラーあり)
        └── ocr_result_no_error.csv  # PDFの文字・数字抽出結果(元の森林調査票にエラーなし)
```

</details>

---

## 7．処理内容

### 7-1．処理内容の概要

まずはDifyで入力ファイルの受付と処理分岐を行い、次にFastAPIを用いたPythonコード処理を実施し、PDFへのOCR処理、森林調査票の入力内容の検査、エラー抽出、GIS用ファイル作成を実行します。
また、LLMを用いて既存ルールでは判定しきれない追加エラー候補を抽出いたします。
最後に、エラー内容についてSlackへ確認内容を通知します。

---

### 7-2．各処理段階の概要

本ツールの、処理段階とその内容、使用した技術の概要は以下の通りです。

| 処理段階 | 内容 | 主な使用技術 |
|---|---|---|
| 1. 入力 | 森林調査票のxlsxまたはPDFをアップロード | Dify |
| 2. ファイル形式判定 | xlsxとPDFで処理を分岐 | Dify |
| 3. PDF OCR | PDFを画像化し、OCRで文字を抽出 | PyMuPDF / Google Cloud Vision API |
| 4. 構造化 | OCR結果をLLMでJSON形式に整理 | Dify / LLM |
| 5. Excel様式変換 | JSONを森林調査票のExcel形式へ変換 | Python / openpyxl / FastAPI |
| 6. 検査 | 調査項目ごとに入力結果を検査 | Python / openpyxl |
| 7. エラー時処理 | エラーログCSVを出力し、修正指示をSlack通知 | CSV / Slack |
| 8. 正常時処理 | GIS用ファイルを出力 | Python / CSV / GIS |
| 9. 追加確認 | LLMで追加エラー候補を抽出 | Dify / LLM |

---

### 7-3．処理方法

#### 7-3-1．入力結果の検査方法

Pythonを用いた入力結果の検査方法は、以下の図の通りです。
まず、森林調査票に対して林調査票の調査項目のセル位置対応表を用いて各調査項目の入力セル位置やデータ範囲を特定します。
その後は、既存のエラールールを定義したファイルから調査項目ごとのルールを取り出し、各調査項目ごとに検査を行っていきます。
最後に、エラーの有無に応じてエラーログ、毎木情報を集計してGIS用ファイルを作成し出力いたします。

[![森林調査票の入力結果の検査方法](https://github.com/yuu-gis0745/prototype/raw/main/docs/images/Python_survey_check.png)](/yuu-gis0745/prototype/blob/main/docs/images/Python_survey_check.png)

**森林調査票の入力結果の検査方法**(画像生成：ChatGPT)

---

#### 7-3-2．Difyによる一括処理

本ツールの、Difyにおける処理のワークフローは以下の通りです。
Dify上で、xlsxもしくはPDF形式の入力結果をアップロードすることで、入力結果の検査～ファイル作成までの処理を一括で行えるワークフローを実装いたしました。

<details>
<summary>ワークフロー図を表示</summary>

```mermaid
flowchart TD;
   A[森林調査票をアップロード<br/>xlsx / PDF] --> B{ファイル形式を判定}

   B -->|xlsx| C[Excel調査票を<br/>そのまま検査対象にする]

   B -->|PDF| D[PDFをOCR処理<br/> PyMuPDF + Google Cloud Vision API]
   D --> E[OCR結果をLLMで<br/>構造化JSONに変換]
   E --> F[JSONをExcel調査票形式へ復元]
   E --> S[OCRのテキスト抽出結果を出力]

   C --> G[FastAPI /analyze で<br/>既存ルール検査を実行]
   F --> G

   G --> H[LLM<br/>追加エラー候補を抽出]
   H --> I[既存エラーと<br/>追加エラー候補を整理]
   I --> J{既存エラーがあったか?}

   J -->|あり| K[error_log.csvを出力]
   K --> M[Slackへ通知]

   J -->|なし| L[GIS用ファイル作成・出力]
   L --> K

   M --> N{エラーがあったか?}

   N -->|既存エラー<br/>追加エラー候補あり| O[既存エラーの修正指示+<br/>追加エラー候補確認指示]
   O --> P[追加エラー候補の<br>既存エラールールファイルへの追加<br/>現在未実装]

   N -->|追加エラー候補あり| Q[追加エラー候補確認指示]

   N -->|いずれのエラーもなし| R[エラーなしの報告]
```

**Difyにおける処理のワークフロー**

</details>

---

## 8．実行方法

このツールを実行するための手順を以下に示します。

---

### 8-1．ローカルでの実行方法

以下のコマンドは、OSはWindows、ターミナルはPowerShellで実行することを想定しております。

```
1. リポジトリをクローン
git clone https://github.com/ユーザー名/リポジトリ名.git

2. 仮想環境を作成・有効化
python -m venv .venv
.\.venv\Scripts\Activate.ps1

3. ライブラリをインストール
python -m pip install -r requirements.txt

4. FastAPIを起動
python -m uvicorn api.main:app --reload

5. `/docs` でAPIを確認
以下のURLにアクセスし、動作を確認する。
http://127.0.0.1:8000
http://127.0.0.1:8000/docs
```

---

### 8-2．Renderにおけるデプロイ設定

本ツールは、Renderにデプロイし、DifyのHTTPリクエストノードからFastAPIのエンドポイントを呼び出して処理を実行する構成としました。
Render 上では、以下のような設定で起動しております(詳細な手順は、deploy_render.mdに整理しております)。

| 項目 | 設定 |
|---|---|
| Runtime | Python |
| Build Command | `pip install -r requirements.txt` |
| Start Command | `uvicorn main:app --host 0.0.0.0 --port $PORT` |
| 使用用途 | Dify から API を呼び出すため |

Pythonのバージョンは、リポジトリ直下の`.python-version`ファイル(`3.12`を指定)でローカル環境とRenderの双方を固定し、環境の違いを防いでおります。

---

### 8-3．Difyとの連携

以下の手順で、Difyでワークフローを実行してください。

```
1．DSLファイルのインポート
Difyのスタジオから「アプリを作成する」→ 「DSLファイルのインポート」で、「forest_survey_check_tool.yml」をインポートします。

2．FastAPIのURL設定
Render_API_BASE_URL には、デプロイ済みFastAPIのURLを設定してください。
例：https://your-render-app.onrender.com

3．ファイルのアップロード
最初のノードである「開始(現地調査結果入力)」→ 「ローカルアップロード」からxlsxもしくはPDF形式の森林調査票をアップロードします。

4．処理開始
「実行開始」を行うと、ワークフローに応じた処理が実行されます。

5．ファイル・Slack投稿文の出力
処理が終了すると、csv形式のエラーログ、GIS用ファイルが出力されるので、ファイル名を付けて任意のフォルダに保存いたします。
```

---

## 9．処理結果・画面例

### 9-1．想定される処理フローの変化

本ツールを導入することで想定される処理フローの改善前、改善後を比較すると、以下のようになります。
一部の作業の自動化を実施することで、手作業とミスを減少させ、全体の作業時間の減少につながることが考えられます。

```mermaid
flowchart TB;
    subgraph Before[改善前：手作業中心]
    direction LR;
        B1[森林調査票<br/>確認]
        B2[Excel様式へ<br/>手入力]
        B3[入力内容を<br/>目視確認]
        B4[ミスを修正]
        B5[GIS用ファイルを<br/>手作業で作成]

        B1 --> B2 --> B3 --> B4 --> B5
    end
    subgraph After[改善後：Python・Difyによる支援]
    direction LR;
        A1[Excel/PDFを<br/>アップロード]
        A2[Pythonで<br/>入力内容を自動検査]
        A3[error_log.csvを<br/>自動出力]
        A4[Difyでエラー内容について<br/>文章を作成]
        A5[qgis_plot_summary.csvを<br/>自動出力]

        A1 --> A2 --> A3 --> A4 --> A5
    end
    Before --> After
```

---

### 9-2．主要な処理結果

ここでは、本ツールの中心となる処理結果を示します。Difyの個別ノードの詳細設定や、PDFからのテキスト抽出ログなど、より細かい内容は [`docs/screenshots.md`](docs/screenshots.md) にまとめておりますので、必要に応じてご参照ください。

まず、Difyの主要ノードは以下の通りです。このワークフローの中で、入力結果の検査を行うHTTPリクエスト(analyze実行)ノード、追加エラー候補の抽出を行うLLM(追加エラー候補抽出)ノード、Slack投稿を行うHTTPリクエスト(Slack投稿)ノードを示しております。

- HTTPリクエスト(analyze実行)ノード
main.pyのanalyzeのエンドポイント実行により、既存エラーのルールが記載されたファイルを元に検査を行います。
- LLM(追加エラー候補抽出)ノード
既存エラー以外にも、LLMに既存エラー以外にもエラーと思われるものがないか検査させ、追加エラー候補として抽出させます。
- HTTPリクエスト(Slack投稿)ノード
エラー内容についてLLMが作成した文章をSlackに投稿いたします。

[![Difyの主要ノード](https://github.com/yuu-gis0745/prototype/raw/main/docs/images/dify_key_nodes.png)](/yuu-gis0745/prototype/blob/main/docs/images/dify_key_nodes.png)

次に、出力されたエラーログは以下の通りです。既存エラーファイル(図上段)と検査を行った森林調査票(図中段)、エラーログ(図下段)を比較すると、例えば「天気」の調査項目でエラーが抽出されていることがわかります。

[![エラーログ](https://github.com/yuu-gis0745/prototype/raw/main/docs/images/error_log.png)](/yuu-gis0745/prototype/blob/main/docs/images/error_log.png)

最後に、作成されたGIS用ファイルをQGIS上に展開したものは以下の通りです。背景図には、15.参考資料・使用データにあるように、地理院タイルを引用しております。GISに展開できるGIS用ファイルとなっていることがわかります。

[![GIS上におけるファイル展開](https://github.com/yuu-gis0745/prototype/raw/main/docs/images/gis_check.png)](/yuu-gis0745/prototype/blob/main/docs/images/gis_check.png)

---

### 9-3．Slack通知

Difyから送付されたSlack投稿文(既存エラー・追加エラー候補あり)の結果は、以下の通りです。
それぞれ担当者に確認することを前提にしつつも、修正方法や修正が必要な理由が記載されております。
また、追加エラー候補の内容を確認すると、樹高に対して胸高直径(木の太さ)が大きいという内容を指摘しており、LLMの方でエラー候補を判断して指摘していることがわかります。

[![Slack投稿文](https://github.com/yuu-gis0745/prototype/raw/main/docs/images/slack_report.png)](/yuu-gis0745/prototype/blob/main/docs/images/slack_report.png)

---

## 10．工夫した点

### 10-1．森林調査票

- 本ツールで使用する森林調査票のExcel様式の入力項目や選択肢は、森林調査業務で使用される項目を想定して作成いたしました。

- 調査項目の中で、斜面位置、斜面方位、異常区分、被害区分の選択肢については、参考文献を元に、実際の森林調査の多くで適用されると思われる選択肢を設定いたしました。
ただし、この制作物で使用しているExcel様式は、卒業制作物のテスト用に私の方で作成したサンプル様式であり実際の森林調査票とは異なっております。

---

### 10-2．入力ファイルの形式

- 本ツールでは、森林調査票のファイル形式をxlsxとPDFといたしました。

- 理由は、調査結果をxlsxへ入力する場合が多かったこと、紙の調査結果もPDFで客先に納品を求められていたことがございます。
PDFから調査結果を直接Excel様式へ精度良く変換・出力することが可能になれば、手入力作業・確認の時間を大幅に短縮できると考え、入力ファイルにPDFも対応するように設計いたしました。

---

### 10-3．セキュリティの設定

- Google Cloud Vision APIの認証JSONは、秘密情報のためGitHubには含めずローカル環境では環境変数 GOOGLE_APPLICATION_CREDENTIALS で管理しております。

- Render環境ではSecret Fileとして登録し、同じ環境変数名から参照する構成にいたしました。

- Slackへの投稿では、Webhook URLはシークレット情報であるためDifyのSecret型環境変数として管理し、GitHubには公開しておりません。

---

### 10-4．その他の設定

- FastAPIを用いてRenderでデプロイすることで、Pythonコード内で実装した入力内容の検査、ファイル出力を実行するように設定いたしました。

- 既存エラーにない新規のエラー候補の抽出を、LLMを用いて行うシステムを実装いたしました。

- Slack通知機能は、エラーチェック後の結果共有を想定して実装しました。

---

## 11．生成AIに支援してもらった部分・自分で担当した部分

### 11-1．生成AIに支援してもらった部分

- Pythonのコード開発、テスト実施、エラー内容の原因特定、コード改善案作成
- FastAPI外部公開方法の検討、実施方法の解説
- Difyワークフローノード内容作成、エラー内容の原因特定、ワークフロー改善案作成
- README構成案作成

### 11-2．自分で担当・確認した部分

**設計・実装で担当した部分**
- 業務改善のための課題の設定
- 森林調査票のExcel様式の作成
- PDFライブラリの選択
- LLMプロンプトの修正

**動作確認で担当した部分**
- 開発されたコードの意味確認
- FastAPIでの動作確認
- DifyとのHTTP連携確認
- LLMのプロンプト内容確認
- 出力CSVの内容確認
- QGISにおけるGIS用ファイル確認

---

## 12．現時点における成果

- 森林調査票の入力内容のチェックの自動化

- PDF調査票からのOCR抽出とExcel形式への出力実装

- エラーログ出力によるエラー内容の一覧化

- GISで利用するためのファイル作成の自動化

- 修正指示や確認事項のSlack通知の実装

---

## 13．現時点における課題

### main.pyの分割化

現時点では、メインの処理も含めて全てmain.pyにコードを記載しており、エラー原因の特定、コード改善の作業が困難となっております。
よって、今後は各エンドポイントごとにコードを分割する必要がございます。

### エラー判別の条件の定数化

既存エラーを判定するときの条件に文字列が設定されており、1文字変更しただけで処理が大きく変化する危険性があります。
よって、今後は文字列の条件の定数化を行う必要がございます。

### 追加エラー候補のファイルへの追加方法

main.pyに追加エラー候補を既存エラールールのファイルに追加するコードを記載しておりますが、現時点のDifyのワークフローでは既存エラールールのファイルへの追加が実装できておりません。
追加エラー候補のファイルへの出力方法は、既存エラールールのファイルに新規追加する方法、別途新規ファイルを作成する方法が考えられますが現在追加方法については検討中です。

### PDFからの文字、数字抽出精度の向上

PyMuPDFとGoogle Cloud Vision APIを組み合わせたPDFからの文字・数字の抽出は、十分な精度を得られませんでした。
原因としては、OCR結果を表形式として復元する処理が不足しており、空欄・重複する「なし」・意図的な入力ミスを精度良く抽出できなかったためと考えられます。
そのため、今後はOCRで読んだ文字の位置を取得する技術の使用が必要であると考えられます。

---

## 14．今後の改善点・取り入れたい技術

- AWS等のクラウドサーバーを用いたFastAPIの実施

- OpenCV等を用いたOCR抽出精度の向上

- RAGを用いたLLMによる追加エラー候補抽出の精度向上

- Shapefile、Geojsonへのファイルフォーマットへの変換

---

## 15．参考資料・使用データ

### 15-1．参考資料

- 林野庁. [森林生態系多様性基礎調査 調査方法の概要（参考）](https://www.rinya.maff.go.jp/j/keikaku/tayouseichousa/attach/pdf/naiyou-6.pdf).
森林調査票の項目設計および調査内容の参考資料として参照。
参照日: 2026-05-17.

### 15-2．使用データ・背景図

- 国土地理院. [地理院タイル一覧](https://maps.gsi.go.jp/development/ichiran.html).
QGISでGISデータを確認する際の背景図として使用。
参照日: 2026-06-27.
