# SimpleExcelXml
OpenXmlを使用しExcelをシンプルな機能のみで簡易的に操作するライブラリを作成しました。

シンプルな機能のみをコンセプトとしているため、セルの書式設定等は実装しておりません(実装することは可能)。

このライブラリはDocumentFormat.OpenXmlを使用して作成されています。

## -説明- Description
・Excelをシンプルな機能のみで作成するインターフェースを実装しています。

・実装している機能は以下の通りになります。

#### 実装(簡易化)されている機能一覧

###### ファイル操作 - File -
・Excelファイルの新規作成

・テンプレートとなるExcelファイルを基にExcelファイルの作成

・保存(Stremを経由した保存)

＊注意

・新規作成の場合、シートが存在しないため手動でシートを追加してください。

###### シート操作 - Sheet -
・シートの追加

・シートの削除

・シート名変更

・データを書き込むシートの切替

###### セル操作 - Cell -
・セルへの書込み(xy座標指定、セル指定(A1等))

・セルの値取得

・行のコピー&ペースト(行の挿入または上書き)

・行の挿入

・オブジェクトを使用したデータの書込み(ExcelSampleObject.csを参照)

#### サンプルファイル
・Program.cs … 動作確認用のプログラムです。実装している機能一覧に記載した機能が動いています。

・ExcelSampleObject.cs … オブジェクトを使用したデータの書込みに使用するサンプルクラスです。

・template.xlsx … ファイルを作成するときにテンプレートとして使用します。

## 豆知識 -Tips-
Excelファイルの拡張子をzipに変更し解凍することができます。

解凍されたファイルの中に「xml」ファイルが含まれており、これをブラウザ上で開くとExcelがどのように構成されているかわかります。

## -必須- Requirement
・[DocumentFormat.OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml/)

## ライセンス - Licence -

・[MIT](https://github.com/yunomichawan/SimpleExcelXml/blob/master/LICENSE)
