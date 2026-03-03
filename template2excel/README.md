# template2excel

## 概要

- このツールは、`Excelファイルを読み込みテンプレートファイル(invoice.schema.json,catalog.schema.json and metadata-def.json)を生成する`ツール(excel2template)の逆の機能を提供します。

- `既存のテンプレートファイル`からテンプレート生成用エクセル書式ファイルを作る際にご利用ください。


## 使い方

- Pythonで実行します
- ファイルは`template2excel.py`のみです
- openpyxlを利用しますのでinstallして利用してください


### helpの表示
```
$ python template2excel.py -h        
usage: template2excel.py [-h] [--dir DIR] [--out OUT] [-v]

options:
  -h, --help       show this help message and exit
  --dir DIR        Specify foldername where template files are stored[default=./template]
  --out OUT        Specify foldername where Excel file will be output [default=./]
  -v, --verbosity  Increase output verbosity
```

読み込むテンプレートファイルをフォルダにまとめて保存し、以下のように実行します。
```
$ python template2excel.py --dir ./template
```

実行すると、ツールを起動したフォルダにRDEDatasetTemplateSheet_template.xlsxが作成されます。なお、読み込み先のフォルダ名がtemplateの場合は以下のようにオプションを省略できます。
```
$ python template2excel.py
```
なお、出力するエクセルファイルは上書きされます。

以上
