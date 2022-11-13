# SRPG Studio 用 ODS コンバーター

ODS ファイルを SRPG Studio にインポート可能な形式のテキストファイルに変換するツールです。

<br>

本ツールを使用することで、SRPG Studio でのシナリオ執筆・イベント作成をより快適に行えるようになります。

<br>

# 使用デモ

Excel や LibreOffice Calc などの表計算ソフトでテンプレート.ods を開き、シート「テンプレート」をもとに指定のフォーマットに沿ってイベント内容を記述します。

<br>

**入力例 1**

![image](https://user-images.githubusercontent.com/117929386/201516394-d2838ee5-4c01-43bc-a6e1-3d4d055cd6d1.PNG)

<br>

**入力例 2**

![image](https://user-images.githubusercontent.com/117929386/201516418-9578f3a5-9dae-4a18-a9c4-f66e671d17b2.PNG)

<br>

ツールを起動し、ODS ファイルを読み込みます。

![image](https://user-images.githubusercontent.com/117929386/201516959-ca4861b9-8bf3-4f28-87f9-8c33b243e2b7.png)

<br>

「全てのシートを変換」は、読み込んだ ODS ファイルの全シートを変換し、1 つのテキストファイルにまとめて出力します。

![image](https://user-images.githubusercontent.com/117929386/201517142-97892812-adbb-44ee-9ca3-62bcb22fe6be.png)

<br>

**変換結果（全シート）**

![image](https://user-images.githubusercontent.com/117929386/201517237-6d63759f-8ace-408d-b7d6-3ffdd073f74b.png)

<br>

「1 つのシートを変換」は、シートを 1 つ選択して変換します。

![image](https://user-images.githubusercontent.com/117929386/201517306-204fd116-40d7-40c6-8263-d033174c0443.png)

<br>

**変換結果（1 シート）**

![image](https://user-images.githubusercontent.com/117929386/201517325-52544940-2fb9-4e89-b55b-c8a1e178dabf.png)

# 開発環境

- Python(3.11.0)
- pandas(1.5.1)
- odfpy(1.4.1)

<br>

# 使い方

## 1.ダウンロード

[リリースノート](https://github.com/sangoopan/srpg-studio-tools/releases/tag/v1.0.0)から SRPG_Studio_ODS.zip をダウンロードします。

![image](https://user-images.githubusercontent.com/117929386/201521178-d42a3f01-36a2-41ea-a0ab-e4981acd923e.png)

<br>

ZIP ファイルを解凍したら、中に以下のファイルが入っていることを確認します。

- srpg_ods_conv.exe
- テンプレート.ods
- readme.txt

![image](https://user-images.githubusercontent.com/117929386/201521267-be87d630-79de-420e-ab68-1159d196d56e.png)

<br>

## 2.テンプレートにイベント内容を入力

テンプレート.ods のシート「テンプレート」をコピー、もしくはそのまま使い、  
以下のフォーマットに沿ってイベント内容を記述します。  
「形式」「位置」「表情」の列はプルダウンリストから内容を選択できます。

<br>

| 形式                           | 発言者                  | 位置       | 表情                                 | 内容                | （備考）                                                                                                                          |
| ------------------------------ | ----------------------- | ---------- | ------------------------------------ | ------------------- | --------------------------------------------------------------------------------------------------------------------------------- |
| コマンド形式                   | ユニット名              | 上/中央/下 | カスタム含む 22 種（空欄は通常扱い） | 表示する文章        |                                                                                                                                   |
| コマンド形式                   | ※                       | ※          | ※                                    | 表示する文章        | 同じコマンド形式&発言者&位置&表情のコマンドが続くとき、※の欄は空欄のままにします。                                                |
| 情報ウィンドウ                 | なし/情報/重要          |            |                                      | 表示する文章        | 情報ウィンドウの種類は発言者の欄で指定します。                                                                                    |
| 選択肢                         | id（半角数字）          |            |                                      | 選択肢 1,選択肢 2,… | 選択肢の id は発言者の欄で指定します。選択肢の内容はカンマ区切りで記述します。                                                    |
| イベント形式（【～イベント】） | イベント id（半角数字） |            |                                      |                     | SRPG Studio にインポートするとき、この行から次の【～イベント】の 1 つ上の行までがひとまとまりのイベントとしてインポートされます。 |
| ※メモ                          |                         |            |                                      |                     | メモの行は変換時に無視されます。                                                                                                  |

<br>

## 3.本ツールで ODS ファイルを読み込み、テキストファイルに変換する

srpg_ods_conv.exe を実行して SRPG Studio 用 ODS コンバーターを起動し、  
2.で編集した ODS ファイルを読み込んで変換します。

テキストファイルは ODS ファイルと同じフォルダに出力され、  
ファイル名は  
「全てのシートを変換」を選択 → "<読み込んだ ODS ファイルの名前>.txt"  
「1 つのシートを変換」を選択 → "<読み込んだ ODS ファイルの名前>\_<選択したシート名>.txt"  
になります。

<br>

## 4.テキストファイルを SRPG Studio にインポートする

3.で変換したテキストファイルを SRPG Studio にインポートします。

<br>

このとき、イベントの種類に合わせたファイル名にする必要があります。  
詳細は[公式サイト](https://srpgstudio.com/lecture/textimport.html "テキストのインポート")をご参照ください。

<br>

# Author

- 作者: さんごぱん
- Twitter: @sangoopan
- Blog: https://sangoopan.hatenablog.com/

# License

Copyright (c) 2022 sangoopan

This software is released under the [MIT license](https://licenses.opensource.jp/MIT/MIT.html).
