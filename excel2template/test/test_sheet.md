excel2templateのテスト用テンプレートシート

## 目的

excel2templateの動作確認をするため以下のテストケースを想定したテンプレートシートを作成した。

テストケースとテスト結果

## Case. 1 送状定義/固有情報の項目なし、試料情報なし

RDEDatasetTemplateSheet_test_01.xlsx

```
test $ python ../excel2template.py RDEDatasetTemplateSheet_test_01.xlsx
RDEDatasetTemplateSheet_test_01.xlsxの処理を開始します。
 - metadata-def.jsonを出力します。
 - invoice.schema.jsonを出力します。
 - invoice.jsonを出力します。
 - catalog.schema.jsonを出力します。
 - catalog.jsonを出力します。
RDEDatasetTemplateSheet_test_01.xlsxの処理を終了します。
Enterを押してください。
```

## Case. 2 送状定義/固有情報の項目あり、試料情報なし

RDEDatasetTemplateSheet_test_02.xlsx

```
test $ python ../excel2template.py RDEDatasetTemplateSheet_test_02.xlsx
RDEDatasetTemplateSheet_test_02.xlsxの処理を開始します。
 - metadata-def.jsonを出力します。
 - invoice.schema.jsonを出力します。
 - invoice.jsonを出力します。
 - catalog.schema.jsonを出力します。
 - catalog.jsonを出力します。
RDEDatasetTemplateSheet_test_02.xlsxの処理を終了します。
Enterを押してください。
```


## Case. 3 送状定義/固有情報の項目あり、試料情報(共通項目あり、一般項目なし、分類別項目なし)

RDEDatasetTemplateSheet_test_03.xlsx

```
test $ python ../excel2template.py RDEDatasetTemplateSheet_test_03.xlsx
RDEDatasetTemplateSheet_test_03.xlsxの処理を開始します。
 - metadata-def.jsonを出力します。
 - invoice.schema.jsonを出力します。
 - invoice.jsonを出力します。
 - catalog.schema.jsonを出力します。
 - catalog.jsonを出力します。
RDEDatasetTemplateSheet_test_03.xlsxの処理を終了します。
Enterを押してください。
```

## Case. 4 送状定義/固有情報の項目あり、試料情報(共通項目あり、一般項目あり、分類別項目なし)

RDEDatasetTemplateSheet_test_04.xlsx

```
test $ python ../excel2template.py RDEDatasetTemplateSheet_test_04.xlsx
RDEDatasetTemplateSheet_test_04.xlsxの処理を開始します。
 - metadata-def.jsonを出力します。
 - invoice.schema.jsonを出力します。
 - invoice.jsonを出力します。
 - catalog.schema.jsonを出力します。
 - catalog.jsonを出力します。
RDEDatasetTemplateSheet_test_04.xlsxの処理を終了します。
Enterを押してください。
```

## Case. 5 送状定義/固有情報の項目あり、試料情報(共通項目あり、一般項目あり、分類別項目あり)

RDEDatasetTemplateSheet_test_05.xlsx

```
test $ python ../excel2template.py RDEDatasetTemplateSheet_test_05.xlsx
RDEDatasetTemplateSheet_test_05.xlsxの処理を開始します。
 - metadata-def.jsonを出力します。
 - invoice.schema.jsonを出力します。
 - invoice.jsonを出力します。
 - catalog.schema.jsonを出力します。
 - catalog.jsonを出力します。
RDEDatasetTemplateSheet_test_05.xlsxの処理を終了します。
Enterを押してください。
```


## Case. 6 送状定義/固有情報の項目あり、試料情報(共通項目あり、一般項目なし、分類別項目あり)

RDEDatasetTemplateSheet_test_06.xlsx

```
test $ python ../excel2template.py RDEDatasetTemplateSheet_test_06.xlsx
RDEDatasetTemplateSheet_test_06.xlsxの処理を開始します。
 - metadata-def.jsonを出力します。
 - invoice.schema.jsonを出力します。
 - invoice.jsonを出力します。
 - catalog.schema.jsonを出力します。
 - catalog.jsonを出力します。
RDEDatasetTemplateSheet_test_06.xlsxの処理を終了します。
Enterを押してください。
```

## Case. 7 送状定義/固有情報の項目あり、試料情報(共通項目なし、一般項目あり、分類別項目なし)

RDEDatasetTemplateSheet_test_07.xlsx

```
test $ python ../excel2template.py RDEDatasetTemplateSheet_test_07.xlsx
RDEDatasetTemplateSheet_test_07.xlsxの処理を開始します。
 - metadata-def.jsonを出力します。
 - invoice.schema.jsonを出力します。
 - invoice.jsonを出力します。
 - catalog.schema.jsonを出力します。
 - catalog.jsonを出力します。
RDEDatasetTemplateSheet_test_07.xlsxの処理を終了します。
Enterを押してください。
```

## Case. 8 送状定義/固有情報の項目あり、試料情報(共通項目なし、一般項目なし、分類別項目あり)

RDEDatasetTemplateSheet_test_08.xlsx

```
test $ python ../excel2template.py RDEDatasetTemplateSheet_test_08.xlsx
RDEDatasetTemplateSheet_test_08.xlsxの処理を開始します。
 - metadata-def.jsonを出力します。
 - invoice.schema.jsonを出力します。
 - invoice.jsonを出力します。
 - catalog.schema.jsonを出力します。
 - catalog.jsonを出力します。
RDEDatasetTemplateSheet_test_08.xlsxの処理を終了します。
Enterを押してください。
```


## Case. 9 送状定義/固有情報の項目なし、試料情報(共通項目あり、一般項目なし、分類別項目なし)

RDEDatasetTemplateSheet_test_09.xlsx

```
test $ python ../excel2template.py RDEDatasetTemplateSheet_test_09.xlsx
RDEDatasetTemplateSheet_test_09.xlsxの処理を開始します。
 - metadata-def.jsonを出力します。
 - invoice.schema.jsonを出力します。
 - invoice.jsonを出力します。
 - catalog.schema.jsonを出力します。
 - catalog.jsonを出力します。
RDEDatasetTemplateSheet_test_09.xlsxの処理を終了します。
Enterを押してください。
```

## Case.10 送状定義/固有情報の項目ありパラメータ名抜け、試料情報なし

RDEDatasetTemplateSheet_test_10.xlsx

```
test $ python ../excel2template.py RDEDatasetTemplateSheet_test_10.xlsx
RDEDatasetTemplateSheet_test_10.xlsxの処理を開始します。
 - metadata-def.jsonを出力します。
 - invoice.schema.jsonを出力します。
 - invoice.jsonの生成に失敗しました。原因: 要件定義（invoice.schema.json）シートのcategory_name='custom'について、必須項目が未指定の行が確認されました: [{'row': 2, 'missing_key': ['parameter_name']}, {'row': 3, 'missing_key': ['label/ja']}, {'row': 5, 'missing_key': ['label/en']}]
 - catalog.schema.jsonを出力します。
 - catalog.jsonを出力します。
RDEDatasetTemplateSheet_test_10.xlsxの処理を終了します。
Enterを押してください。
```

## Case. 11 送状定義/固有情報の項目ありパラメータ名重複、試料情報なし
  - RDEDatasetTemplateSheet_test_11.xlsx
```
test $ python ../excel2template.py RDEDatasetTemplateSheet_test_12.xlsx
RDEDatasetTemplateSheet_test_12.xlsxの処理を開始します。
 - metadata-def.jsonを出力します。
 - invoice.schema.jsonを出力します。
 - invoice.jsonを出力します。
 - catalog.schema.jsonを出力します。
 - catalog.jsonを出力します。
RDEDatasetTemplateSheet_test_12.xlsxの処理を終了します。
Enterを押してください。
```

## Case. 12 送状定義/固有情報の項目あり、試料情報なし、maximum列より右側なし

template2excelの生成物対応

RDEDatasetTemplateSheet_test_12.xlsx

```
test $ python ../excel2template.py RDEDatasetTemplateSheet_test_12.xlsx
RDEDatasetTemplateSheet_test_12.xlsxの処理を開始します。
 - metadata-def.jsonを出力します。
 - invoice.schema.jsonを出力します。
 - invoice.jsonを出力します。
 - catalog.schema.jsonを出力します。
 - catalog.jsonを出力します。
RDEDatasetTemplateSheet_test_12.xlsxの処理を終了します。
Enterを押してください。
```

## Case. 13 カタログ定義/正常系(textareaなど各種データ型含む)

RDEDatasetTemplateSheet_test_13.xlsx

```
test $ python ../excel2template.py RDEDatasetTemplateSheet_test_13.xlsx
RDEDatasetTemplateSheet_test_13.xlsxの処理を開始します。
 - metadata-def.jsonを出力します。
 - invoice.schema.jsonを出力します。
 - invoice.jsonを出力します。
 - catalog.schema.jsonを出力します。
 - catalog.jsonを出力します。
RDEDatasetTemplateSheet_test_13.xlsxの処理を終了します。
Enterを押してください。
```

## Case. 14 カタログ定義/項目重複あり

RDEDatasetTemplateSheet_test_14.xlsx

```
test $ python ../excel2template.py RDEDatasetTemplateSheet_test_14.xlsx
RDEDatasetTemplateSheet_test_14.xlsxの処理を開始します。
 - metadata-def.jsonを出力します。
 - invoice.schema.jsonを出力します。
 - invoice.jsonを出力します。
 - catalog.schema.jsonを出力します。
 - catalog.jsonの生成に失敗しました。原因: 要件定義（catalog.schema.json）シートのcategory_name='parameter_name'について、重複する行が確認されました: {'data_creator'}
RDEDatasetTemplateSheet_test_14.xlsxの処理を終了します。
Enterを押してください。
```


## Case.15 カタログ定義/項目名抜け

RDEDatasetTemplateSheet_test_15.xlsx

```
test $ python ../excel2template.py RDEDatasetTemplateSheet_test_15.xlsx
RDEDatasetTemplateSheet_test_15.xlsxの処理を開始します。
 - metadata-def.jsonを出力します。
 - invoice.schema.jsonを出力します。
 - invoice.jsonを出力します。
 - catalog.schema.jsonを出力します。
 - catalog.jsonの生成に失敗しました。原因: 要件定義（catalog.schema.json）シートのcategory_name='parameter_name'について、必須項目が未指定の行が確認されました: [{'row': 3, 'missing_key': ['parameter_name']}, {'row': 4, 'missing_key': ['label/ja']}, {'row': 5, 'missing_key': ['label/en']}]
RDEDatasetTemplateSheet_test_15.xlsxの処理を終了します。
Enterを押してください。
```

## Case. 16 カタログ定義/options/unitから右列なし

RDEDatasetTemplateSheet_test_16.xlsx

```
test $ python ../excel2template.py RDEDatasetTemplateSheet_test_16.xlsx
RDEDatasetTemplateSheet_test_16.xlsxの処理を開始します。
 - metadata-def.jsonを出力します。
 - invoice.schema.jsonを出力します。
 - invoice.jsonを出力します。
 - catalog.schema.jsonを出力します。
 - catalog.jsonを出力します。
RDEDatasetTemplateSheet_test_16.xlsxの処理を終了します。
Enterを押してください。
```  

## Case. 17 メタデータ定義/正常系
RDEDatasetTemplateSheet_test_17.xlsx

```
test $ python ../excel2template.py RDEDatasetTemplateSheet_test_17.xlsx
RDEDatasetTemplateSheet_test_17.xlsxの処理を開始します。
 - metadata-def.jsonを出力します。
 - invoice.schema.jsonを出力します。
 - invoice.jsonを出力します。
 - catalog.schema.jsonを出力します。
 - catalog.jsonを出力します。
RDEDatasetTemplateSheet_test_17.xlsxの処理を終了します。
Enterを押してください。
```

## Case.18 メタデータ定義/パラメータ名抜け

RDEDatasetTemplateSheet_test_18.xlsx

```
test $ python ../excel2template.py RDEDatasetTemplateSheet_test_18.xlsx
RDEDatasetTemplateSheet_test_18.xlsxの処理を開始します。
 - metadata-def.jsonの生成に失敗しました。原因: 要件定義（metadata-def.json）シートのcategory_name='metadata'について、必須項目が未指定の行が確認されました: [{'row': 2, 'missing_key': ['name/ja']}, {'row': 3, 'missing_key': ['name/en']}]
 - invoice.schema.jsonを出力します。
 - invoice.jsonを出力します。
 - catalog.schema.jsonを出力します。
 - catalog.jsonを出力します。
RDEDatasetTemplateSheet_test_18.xlsxの処理を終了します。
Enterを押してください。
```


## Case.19 メタデータ定義/パラメータ名重複

RDEDatasetTemplateSheet_test_19.xlsx

```
test $ python ../excel2template.py RDEDatasetTemplateSheet_test_19.xlsx
RDEDatasetTemplateSheet_test_19.xlsxの処理を開始します。
 - metadata-def.jsonの生成に失敗しました。原因: 要件定義（metadata-def.json）シートのcategory_name='metadata'について、重複する行が確認されました: {'key_datetime', 'key_number'}
 - invoice.schema.jsonを出力します。
 - invoice.jsonを出力します。
 - catalog.schema.jsonを出力します。
 - catalog.jsonを出力します。
RDEDatasetTemplateSheet_test_19.xlsxの処理を終了します。
Enterを押してください。
```

## Case.20 $id、$schema抜け

RDEDatasetTemplateSheet_test_20.xlsx

```
test $ python ../excel2template.py RDEDatasetTemplateSheet_test_20.xlsx
RDEDatasetTemplateSheet_test_20.xlsxの処理を開始します。
 - metadata-def.jsonを出力します。
 - invoice.schema.jsonの生成に失敗しました。原因: 要件定義（invoice.schema.json）の$schema、$id値は必須です。
 - invoice.jsonを出力します。
 - catalog.schema.jsonを出力します。
 - catalog.jsonを出力します。
RDEDatasetTemplateSheet_test_20.xlsxの処理を終了します。
Enterを押してください。
```



