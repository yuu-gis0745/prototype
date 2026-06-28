# Renderへのデプロイ手順

このドキュメントでは、森林調査データ入力・検査効率化ツールを **Render** にデプロイする際の設定内容を整理します。<br>
README.mdには実行方法の概要だけを記載し、Renderの詳しい設定や注意点は本ファイルに分けて記載します。

---

## 1. 前提となるフォルダ構成

この手順では、以下のフォルダ構成に示すように、FastAPIの `main.py` が `api` フォルダ内にある構成を前提とします。

```
forest_survey_check_tool/
├─ api/
│  └─ main.py
├─ requirements.txt
├─ README.md
└─ docs/
   └─ deploy_render.md
```

この構成では、Uvicornの起動対象は次のように指定します。


```
api.main:app
```

意味は以下の通りです。

| 書き方 | 意味 |
|---|---|
| `api.main` | `api` フォルダ内の `main.py` |
| `app` | `main.py` の中で定義している `FastAPI()` の変数 |


例：

```python
from fastapi import FastAPI

app = FastAPI()
```

---

## 2. ローカル環境での実行方法

Renderへデプロイする前に、ローカル環境で起動できることを確認します。<br>
以下は **Windows PowerShell** での実行例です。

```
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m uvicorn api.main:app --reload
```

起動後、ブラウザで以下にアクセスして動作確認を行います。


| URL | 確認内容 |
|---|---|
| `http://127.0.0.1:8000` | APIのトップページ確認 |
| `http://127.0.0.1:8000/docs` | FastAPIの自動ドキュメント確認 |

---

## 3. Renderでの基本設定

Renderでは、GitHubリポジトリを接続してWebサービスとしてデプロイします。<br>
設定項目の例は以下の通りです。

| 項目 | 設定例 |
|---|---|
| Runtime | Python |
| Build Command | `python -m pip install -r requirements.txt` |
| Start Command | `python -m uvicorn api.main:app --host 0.0.0.0 --port $PORT` |

---

## 4. RenderのStart Command

Renderで使用するStart Commandは以下です。<br>

```
python -m uvicorn api.main:app --host 0.0.0.0 --port $PORT
```

Start Commandのそれぞれの部分は、以下に示す意味を持っております。

| 部分 | 意味 |
|---|---|
| `python -m uvicorn` | 現在のPython環境に入っているUvicornを使って起動する |
| `api.main:app` | `api/main.py` の中の `app` を起動する |
| `--host 0.0.0.0` | 外部からアクセスできるようにする |
| `--port $PORT` | Renderが指定するポート番号を使う |

---

## 5. ローカル実行とRender実行の違い

ローカル環境では、開発中にコード変更を反映しやすくするため `--reload` を付けます。

```
python -m uvicorn api.main:app --reload
```

一方、Renderでは通常 `--reload` は使いません。

```
python -m uvicorn api.main:app --host 0.0.0.0 --port $PORT
```

違いは以下の通りです。

| オプション | 使用場面 | 意味 |
|---|---|---|
| `--reload` | ローカル開発用 | コード変更時に自動再起動する |
| `--host 0.0.0.0` | Render用 | 外部からアクセスできるようにする |
| `--port $PORT` | Render用 | Render側で指定されるポートを使う |

---

## 6. 環境変数・Secret Fileの注意点

APIキー、Slack Webhook URL、Google認証JSONなどの秘密情報は、GitHubやREADMEに直接記載しません。<br>
公開してはいけない情報の例は以下です。

- APIキー
- Slack Webhook URL
- Google認証JSONの中身
- Secret Fileの中身
- 個人用の一時URL
- 認証情報を含むURL

Renderで秘密情報を扱う場合は、Renderの管理画面で環境変数やSecret Fileとして設定します。<br>
READMEや本ファイルには、**設定が必要であることだけ** を書き、実際の値は記載しないようにします。

---

## 7. Difyとの連携

本ツールは、DifyのHTTP RequestノードからRenderにデプロイしたFastAPIのエンドポイントを呼び出す構成です。<br>
基本的な流れは以下の通りです。

```
Dify
  ↓ HTTP Request
Render上のFastAPI
  ↓
検査処理・CSV作成など
  ↓
Difyへ結果を返す
```

Dify側では、Renderで発行されたURLをHTTPリクエストノードのURLに設定します。

例：

```
https://xxxxx.onrender.com/エンドポイント名
```

RenderのURLを複数のHTTPリクエストノードで使う場合は、転記ミスを防ぐため、Dify側で環境変数として管理する方法も検討します。

---

## 8. 動作確認の流れ

Renderにデプロイした後は、以下の順番で確認します。<br>

### 8.1 トップページの確認

Renderで発行されたURLにアクセスします。<br>

```
https://xxxxx.onrender.com
```

`{"status": "ok"}` のようなレスポンスが返る場合、FastAPI自体は起動できています。

### 8.2 FastAPIドキュメントの確認

以下にアクセスします。

```
https://xxxxx.onrender.com/docs
```

FastAPIの自動ドキュメントが表示されれば、エンドポイントの確認やテストができます。

### 8.3 Difyからの呼び出し確認

DifyのHTTPリクエストノードから、Render上のAPIを呼び出します。<br>
確認する内容は以下です。

- URLが正しいか
- HTTPメソッドが正しいか
- Headersの設定が正しいか
- BodyのJSON形式が正しいか
- Render側でエラーが出ていないか
- Dify側でレスポンスを受け取れているか

---

## 9. よくあるエラーと確認ポイント

### 9.1 `requirements.txt` が見つからない

原因としては、requirements.txtの格納場所に問題があることが多いです。

エラー例：

```text
ERROR: Could not open requirements file: [Errno 2] No such file or directory: 'requirements.txt'
```

確認すること：

- `requirements.txt` がリポジトリ直下にあるか
- RenderのBuild Commandがリポジトリ直下を前提にしているか
- GitHubに `requirements.txt` をpushしているか

---

### 9.2 `main` が読み込めない

原因として、`main.py` の場所とUvicornの指定が合っていない可能性があります。

`main.py` がリポジトリ直下にある場合：

```
python -m uvicorn main:app --reload
```

`main.py` が `api` フォルダ内にある場合：

```
python -m uvicorn api.main:app --reload
```

現在の構成では `api/main.py` のため、Renderでも以下を使います。

```
python -m uvicorn api.main:app --host 0.0.0.0 --port $PORT
```

---

### 9.3 Renderでは起動するが、Difyから呼び出せない

新しくRenderのURLを発行したときに発生しやすいエラーです。<br>原因としては、DifyのHTTPリクエストノードのURLを古いURLのままにしていることや、URLの転記ミスが考えられます。

確認すること：

- DifyのHTTPリクエストノードのURLが正しいか
- エンドポイント名が正しいか
- POST / GET などのHTTPメソッドが合っているか
- `Content-Type: application/json` が必要な場合、Headersに設定しているか
- Bodyが正しいJSON形式になっているか
- Renderのログにエラーが出ていないか

---

### 9.4 Secret Fileが見つからない

Secret Fileを使っている場合、Render側でファイルが正しく設定されていないとエラーになります。

確認すること：

- RenderにSecret Fileを保存しているか
- ファイル名や参照パスがコードと一致しているか
- 再デプロイが必要な場合には、認証設定を保存後に再デプロイしているか
- GitHubに認証JSONを直接置いていないか

---

## 10. 最終確認チェックリスト

デプロイ前後に、以下を確認します。

| 確認項目 | 確認 |
|---|:---:|
| `requirements.txt` がリポジトリ直下にある |  |
| `api/main.py` に `app = FastAPI()` がある |  |
| ローカルで `python -m uvicorn api.main:app --reload` が成功する |  |
| RenderのBuild Commandが設定されている |  |
| RenderのStart Commandが正しい |  |
| 秘密情報をGitHubにpushしていない |  |
| 必要な環境変数やSecret FileをRenderに設定している |  |
| RenderのURLでトップページを確認した |  |
| RenderのURL + `/docs` を確認した |  |
| DifyのHTTPリクエストノードから呼び出せる |  |

---

## 11. まとめ

ローカル環境とRenderでは、Uvicornの起動コマンドが異なります。

ローカル環境：

```
python -m uvicorn api.main:app --reload
```

Render：

```
python -m uvicorn api.main:app --host 0.0.0.0 --port $PORT
```

Renderでは、外部からアクセスできるように `--host 0.0.0.0` を指定し、ポート番号はRenderが指定する `$PORT` を使います。<br>
また、APIキーや認証JSONなどの秘密情報はGitHubやREADMEに記載せず、Renderの環境変数やSecret Fileで管理します。
