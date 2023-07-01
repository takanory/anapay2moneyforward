# anapay2moneyforward

* ANA Payのメールから支払い情報を取り出して、マネーフォワードに登録するスクリプト。

## 環境構築

* 必要なライブラリをvenvにインストールする

```bash
$ python 3.11 -m venv env
$ . env/bin/activate
(env) $ pip install -r reqirements.txt
```

## GmailとGoogle SpreadsheetのAPIを使えるようにする

* 以下のページも参考にして、GoogleのAPIを使えるようにする
  * [Python クイックスタート  |  Gmail  |  Google for Developers](https://developers.google.com/gmail/api/quickstart/python?hl=ja)
* Google Cloudコンソールでプロジェクトを作成する
  * プロジェクトでGmail APIとGoogle Sheets APIを有効にする
  * OAuth 同意画面でアプリを作成する
  * テストユーザーで自分のGoogleアカウントを追加
  * 認証情報をダウンロードし、`credentials.json` として保存
* 以下のように `quickstart.json` を実行する
  * 自分のGoogleアカウントで **同意する**
  * 処理が成功すると `token.json` が生成される

```
(env) $ python quickstart.py
(env) $ ls *.json
credentials.json	token.json
```
