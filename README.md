# RDE/データセットテンプレート生成、確認ツール

## RDEおよび提供するツールについて

RDE (Research Data Express) は、物質・材料についての研究データをオンラインで迅速に登録するため国立研究開発法人物質・材料研究機構（以下、NIMS）が開発したシステムです。生データを登録すると自動的にデータ駆動型のマテリアル研究に適した形に構造化してクラウドに蓄積します。これによりユーザーや研究グループ内での再利用や他の研究グループとのデータの共用が容易となり、マテリアル研究開発のDX化を支援します。

ここで提供するツールは、RDEにおけるデータセットテンプレート開発を補助するためのツールです。

以下の機能を提供しています。
- RDEのデータセットを開設、運用するためのデータセットテンプレートを構成するファイルを作成、プレビューするためのツールです
- データセットテンプレートのうち、送り状スキーマファイル(invoice.schema.json)、メタデータ定義(metadata-def.json)を作成するとツールと、その結果をVScodeでプレビューするツールの２つで構成されます
- エクセル形式のファイルに必要事項を入力することでjson形式のファイルを出力することができます
- また、既存のテンプレートファイルからテンプレートを生成するためのエクセル書式ファイルを作成するツールもあります(template2excel参照)
<br />

## 利用方法

  docsフォルダ内のファイルにてご確認ください。

  template2excelについては[こちら](./template2excel/README.md)を参照してください。

<br />


## 動作環境

* Windows実行系　Windows 10以上(pyintallerで作成)
* Mac OS実行系(arm64)(pyintallerで作成)
* pythonコードの実行確認はPython 3.13以上
* template2excelについては実行系は配布していません

実行系の作成については[こちら](./excel2template/pyinstaller_build.md)を参照してください。

<br />

## 操作方法

* 取扱説明書を参考にご利用ください。
* Windowsにて利用する場合は、excel2template/dist/excel2template.exeを利用してください。
* Macにて利用する場合は、excel2template/dist/excel2templateを利用してください。
* pythonのコードを利用して実行する場合は、excel2template/excel2template.pyを取得してください。
* VScodeの追加機能はtemplate_viewerからtemplate-viewer-1.0.0.vsixを取得してください。
* excel2templateのテストには、[RDEDatasetTemplateSheet_20251222_sample.xlsx](./excel2template/RDEDatasetTemplateSheet_20251222_sample.xlsx)を利用してください。生成結果は[こちら](./excel2template/RDEDatasetTemplateSheet_20251222_sample/)に掲載しています。

<br />

## 利用ルールおよびライセンス
 
* 本プログラムはMITライセンスで提供されています。

<br />

### RDEおよび本ツール関するお問い合わせ

ご不明な点につきましては、以下にお問い合わせください。

<br />

国立研究開発法人物質・材料研究機構
技術開発・共用部門 材料データプラットフォーム

お問い合わせ フォーム<br>
https://dice.nims.go.jp/contact/form.html

<br />

2025.5.15 公開
