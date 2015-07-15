# xls_renderer

## 概要

.xls/.xlsx形式のテンプレートファイルに.csv形式のデータを差し込みして.xls/.xlsx形式を出力します。

## 実行例

```shell
xls_renderer -d data.csv -p parameter.csv -pc 8 -pr 8 -t template.xls -o output.xls

xls_renderer -d data.csv -p parameter.csv -pc 8 -pr 8 -t template.xlsx -o output.xlsx -x
```

## コマンドラインオプション

### -d data.csv

* データ用の.csvファイル名を指定します。
* 1行目がテンプレートファイル中で置換対象となるキー名となります。
* テンプレートファイル中のセルの値が「#」または「＃」＋キー名のものが2行目の値で置換されます。
* 画像ファイルを挿入する場合はファイルの絶対パスを値としてください。

### -p parameter.csv

* パラメータ用のCSVファイル名を指定します。
* 1行目がテンプレートファイル中で置換対象となるキー名となります。
* 2行目がセルの値のタイプを表します。
 * string: 文字列
 * numeric: 数値
 * formula: 数式
 * image_jpg,w,h: JPEG形式の画像ファイル、wとhには幅と高さをセル数で指定します。
 * image_png,w,h: PNG形式の画像ファイル、wとhには幅と高さをセル数で指定します。

### -pc 8

* 画像ファイルを挿入する場合の1列のピクセル数。
* 省略時のデフォルト値は8です。

### -pr 8

* 画像ファイルを挿入する場合の1行のピクセル数。
* 省略時のデフォルト値は8です。

### -t template.xls

* テンプレート用の.xls/xlsxファイル名を指定します。

### -o output.xls

* 出力する.xls/xlsxファイル名を指定します。

### -x

* テンプレートと出力のファイル形式を.xlsxとして扱います。指定しない場合は.xlsとして扱います。
