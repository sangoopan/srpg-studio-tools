# SRPG Studio 用テキストコンバーター

Excel で編集したシナリオ（イベント内容や世界観設定など）を SRPG Studio にインポート可能な形式のテキストファイルに変換するツールです。

<br>

本ツールを使用することで、SRPG Studio でのシナリオ執筆・イベント作成をより快適に行えるようになります。

<br>

# 使用デモ

同梱のテンプレート.xlsx を Excel や LibreOffice Calc などの表計算ソフトで開き、シート「テンプレート」をもとに指定のフォーマットに沿ってイベント内容などを記述します。

<br>

**入力例 1**

![image](https://user-images.githubusercontent.com/117929386/201516394-d2838ee5-4c01-43bc-a6e1-3d4d055cd6d1.PNG)

<br>

**入力例 2**

![image](https://user-images.githubusercontent.com/117929386/201516418-9578f3a5-9dae-4a18-a9c4-f66e671d17b2.PNG)

<br>

ツールを起動し、xlsx ファイルを読み込みます。

![image](https://user-images.githubusercontent.com/117929386/202430619-1a1fad2f-1721-4f0f-a9bc-7d4fbd985bb6.png)

<br>

ファイルはドラッグ&ドロップで開くこともできます。

![gif](https://user-images.githubusercontent.com/117929386/202434694-280569eb-8178-4212-a008-bb896c4a742a.gif)

<br>

「全てのシートを変換」は、読み込んだ xlsx ファイルの全シートを変換し、1 つのテキストファイルにまとめて出力します。

![image](https://user-images.githubusercontent.com/117929386/202430732-c276eb6c-5fe4-4717-be5b-231ffdd257c0.png)

<br>

**変換結果（全シート）**

![image](https://user-images.githubusercontent.com/117929386/201517237-6d63759f-8ace-408d-b7d6-3ffdd073f74b.png)

<br>

「1 つのシートを変換」は、シートを 1 つ選択して変換します。

![image](https://user-images.githubusercontent.com/117929386/202430808-9eea16c4-160a-4c10-bf87-338920e59d63.png)

<br>

**変換結果（1 シート）**

![image](https://user-images.githubusercontent.com/117929386/201517325-52544940-2fb9-4e89-b55b-c8a1e178dabf.png)

# 開発環境

言語

- Python(3.11.0)

<br>

ライブラリ

- openpyxl(3.0.10)
- TkinterDnD2(0.3.0)

<br>

# 使い方

## 1.ダウンロード

[リリースノート](https://github.com/sangoopan/srpg-studio-tools/releases/tag/v2.0.0)から SRPG_Studio_Conv.zip をダウンロードします。

![image](https://user-images.githubusercontent.com/117929386/202439794-7917e85d-eb85-41cf-9bde-b91e1e7ffd54.png)

<br>

ZIP ファイルを解凍したら、中に以下のファイルが入っていることを確認します。

- srpg_text_conv.exe
- テンプレート.xlsx
- readme.txt

![image](https://user-images.githubusercontent.com/117929386/202431659-80564d3b-3b76-468f-99c3-7f162eaebc65.png)

<br>

## 2.テンプレートにイベント内容を入力

テンプレート.xlsx のシート「テンプレート」をコピー、もしくはそのまま使い、  
以下のフォーマットに沿ってイベント内容を記述します。  
「形式」「位置」「表情」の列はプルダウンリストから内容を選択できます。

<br>

| 形式                           | id          | 名前                   | 位置        | 表情                | 内容                   |
| :----------------------------- | :---------- | :--------------------- | :---------- | :------------------ | :--------------------- |
| メッセージ                     |             | ユニット名※            | 上/中央/下※ | カスタム含む 22 種※ | 表示する文章           |
| テロップ                       |             |                        | 上/中央/下※ |                     | 表示する文章           |
| メッセージタイトル             |             |                        | 上/中央/下※ |                     | 表示する文章           |
| スチルメッセージ               |             | 名前※                  |             |                     | 表示する文章           |
| 情報ウィンドウ                 |             | なし/情報/重要※        |             |                     | 表示する文章           |
| メッセージスクロール           |             |                        | 上/中央/下※ |                     | 表示する文章           |
| 選択肢                         | 選択肢の ID |                        |             |                     | 選択肢（カンマ区切り） |
| 【ユニットイベント】           | ユニット ID | ユニット名@イベント ID |             |                     |                        |
| 【場所イベント】               | イベント ID |                        |             |                     |                        |
| 【自動開始イベント】           | イベント ID |                        |             |                     |                        |
| 【会話イベント】               | イベント ID |                        |             |                     |                        |
| 【オープニングイベント】       | イベント ID |                        |             |                     |                        |
| 【エンディングイベント】       | イベント ID |                        |             |                     |                        |
| 【コミュニケーションイベント】 | イベント ID |                        |             |                     |                        |
| 【回想イベント】               | イベント ID |                        |             |                     |                        |
| 【マップ共有イベント】         | イベント ID |                        |             |                     |                        |
| 【ブックマークイベント】       | イベント ID |                        |             |                     |                        |
| 【世界観設定】                 |             | 名前                   |             |                     |                        |
| \#メモ                         |             |                        |             |                     |                        |

<br>

**特記事項**

- ※印がついている欄は、上の行と同じ「形式」で、且つその行内の※欄の変更が必要ないとき空欄にします。
- 【～イベント】は、その行から次の【～イベント】の 1 つ上の行までが同じイベントとして扱われます。SRPG Studio にインポートする際、この記述がないと正常に反映されない可能性があります。
- #メモの行は変換結果に反映されません。イベント演出のメモや ToDo などご自由にお使いください。

<br>

テンプレート.xlsx のシート「入力例 1」「入力例 2」「入力例 3」に入力例が載っているので、適宜ご参照ください。

<br>

## 3.本ツールで xlsx ファイルを読み込み、テキストファイルに変換する

srpg_text_conv.exe を実行して SRPG Studio 用テキストコンバーターを起動し、  
2.で編集した xlsx ファイルを読み込んで変換します。

テキストファイルは変換元の xlsx ファイルと同じフォルダに出力され、  
ファイル名は  
「全てのシートを変換」を選択 → "<読み込んだ xlsx ファイルの名前>.txt"  
「1 つのシートを変換」を選択 → "<読み込んだ xlsx ファイルの名前>\_<選択したシート名>.txt"  
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
