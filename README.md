
# 1. 事前準備

## 1.1 Python のインストール

### 1.1.1 Windows + R キーを押下し「ファイル名を指定して実行」画面を表示し cmd と入力して OK します

<img width="286" alt="image" src="https://github.com/flatrock/rss_assist/assets/7097104/5de64461-30d4-4d61-8849-2bcdc2d75301">

### 1.1.2 コマンドプロンプトで python と入力して Enter します

<img width="867" alt="image" src="https://github.com/flatrock/rss_assist/assets/7097104/b52ca1ac-f11a-4892-a882-fdbd305dac26">

### 1.1.3 Microsoft Store 画面で「入手」ボタンをクリックして Python をインストールします

<img width="902" alt="image" src="https://github.com/flatrock/rss_assist/assets/7097104/c99481ed-f170-4b32-9fe7-1828b51ec329">

## 1.2 xlwings のインストール

- コマンドプロンプトで pip install xlwings と入力して Enter します

<img width="867" alt="image" src="https://github.com/flatrock/rss_assist/assets/7097104/665b7e1c-92e1-4229-8eba-643829e7c9a3">

# 2. 実行

## 2.1 Market Speed II アプリの起動

- Market Speed II アプリを起動します

## 2.2 Excel ブックから Market Speed II サーバーへの接続

- Excel ブック (.xlsx) を開き、「マーケットスピード II」リボンを接続中の状態にします

<img width="537" alt="image" src="https://github.com/flatrock/rss_assist/assets/7097104/17e61ef3-b356-4548-b6b0-7803e479dd7b">

## 2.3 Python スクリプトの実行

- 2.3.1 [rss_assist.py](https://raw.githubusercontent.com/flatrock/rss_assist/main/rss_assist.py) をダウンロードし任意のフォルダーに配置します

- 2.3.2 rss_assist.py を配置したフォルダーをエクスプローラーで開きアドレスバーに cmd と入力して Enter します

<img width="585" alt="image" src="https://github.com/flatrock/rss_assist/assets/7097104/d6262804-11d1-48ca-aa79-8f76c5eca422">

- 2.3.3 コマンドプロンプトで python rss_assist.py と入力し Enter します
<img width="867" alt="image" src="https://github.com/flatrock/rss_assist/assets/7097104/6664ee4b-dac5-4ac6-83ff-9e8373e62053">

- 2.3.4 rss_assist.py の実行中は注目値 (デフォルトで現在値と売買圧力比率) の変化方向に応じてセルの背景色が変化します（増加は青、減少は赤）
<img width="1051" alt="image" src="https://github.com/flatrock/rss_assist/assets/7097104/18794759-adbc-4273-82b8-af1b3fac25ec">

- 2.3.5 rss_assist.py を終了するには Ctrl + C キーを押下するか、コマンドプロンプトを閉じます
<img width="867" alt="image" src="https://github.com/flatrock/rss_assist/assets/7097104/8459923f-240e-4349-93f2-f4857dcee850">
