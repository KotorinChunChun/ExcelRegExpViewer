# [ExcelRegExpViewer - Excel正規表現確認ビューアー](https://github.com/KotorinChunChun/ExcelRegExpViewer/)

## これは

Excelのテーブル上で正規表現の動作確認をしつつ、ノウハウを蓄積していくためのエクセルマクロブックです。

正規表現には多種多様な方言があります。このツールは、 Excel VBAなどで使われる `VBScript.RegExp` のライブラリのサポートしている記法しか対応していないので注意が必要です。

巷には正規表現をグラフィカルに検証する便利なWEBサイトが多数あるものの、事例を蓄積するのには適しているとは言えないため、このようなものを開発しました。

ワークシート上でVBAのユーザー定義関数を使って正規表現を検証し、HYPERLINK・ENCODEURL関数を使って外部WEBサイトにハイパーリンクしてグラフィカルなプレビューを丸投げしています。

## できること

* Excelワークシート上で正規表現の動作確認
* 作成したパターンをワンクリックで https://regexper.com/ に転送し、フローチャートでプレビュー
* VBA、特にイミディエイトウィンドウから手っ取り早く正規表現を実行

![image](https://user-images.githubusercontent.com/55196383/93624710-ba025f00-fa1b-11ea-80a2-f9de16690859.png)

![image](https://user-images.githubusercontent.com/55196383/93624577-8b848400-fa1b-11ea-9099-c7430d099133.png)

![image](https://user-images.githubusercontent.com/55196383/93623168-40697180-fa19-11ea-9c5b-30bbfb75b755.png)

![image](https://user-images.githubusercontent.com/55196383/93623410-aa821680-fa19-11ea-9918-9d0b82c4ffb4.png)

## シート

### 正規表現テスト

テーブル形式でソース、パターン、意味を入力すると、UDFで実行結果を右に表示してくれる

### 正規表現SubMatchテスト

ソース、パターン、意味を入力すると、UDFで実行結果を二次元表示してくれる

## 関数一覧

#### RegexIsMatch - マッチするかを確認

#### RegexReplace - マッチした文字列を置換

#### RegexMatchCount - マッチした箇所の個数

#### RegexMatchIndexs - マッチした箇所の開始インデックス配列

#### RegexMatchLengths - マッチした箇所の文字列長配列

#### RegexMatchValues - マッチした箇所の値配列

#### RegexSubMatches - マッチした箇所の配列のサブマッチ値配列

## その他

細かいところは各自でカスタマイズすべし。

良いアイディア会ったらフィードバックよろしくぅ

## 参考 - 正規表現チェックツールまとめ

https://qiita.com/aqril_1132/items/c185c7ad84c129e5a2df#1debuggex-online-visual-regex-tester

## The MIT License

Copyright (c) 2020 [KotorinChunChun](https://github.com/KotorinChunChun)


