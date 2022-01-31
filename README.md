# Excel+PythonでArcGIS Online / Portal のアイテム詳細を一括更新
Excel × ArcGIS API for Python でArcGIS Online（もしくはPortal）のアイテム詳細を一括更新するサンプルです。

## 概要

Excel に整理してあるArcGIS Online（もしくはPortal） のアイテム詳細の情報を、ArcGIS API for Python を使って、一括更新するスクリプトとExcel のサンプルです。
同時にこのスクリプト内では、アイテムの削除防止の有効化、特定グループに共有、所有者の変更、パブリックで公開するまでを行っています。

2021年に[ArcGIS Hub Basic](https://www.esrij.com/products/arcgis-hub/) の機能を使って作成した「[3D 都市モデルの ArcGIS における活用事例を集めた活用サイト](https://3d-city-model.esrij.com/)」に掲載したファイル ジオデータベース (FGDB) のアイテム詳細の更新時にこのスクリプトを利用しました。

特定グループは任意にidで指定可能ですが、今回のサンプルでは、ESRIジャパンのArcGIS Online 組織サイト（https://ej.maps.arcgis.com/ ）の「国土交通省 3D都市モデル 「Project PLATEAU」の活用 のコンテンツ」(id:d568e2190c4f456f8bd812c8e07c719a) のグループに共有することで、ArcGIS Hub 側で情報の登録を重複して行うことことなく、最終的には、[カタログサイト](https://3d-city-model.esrij.com/search?collection=Document&sort=name) から、*105* のアイテムを公開した状態になっております。

Excel の整理例は、[3D都市モデル_AGOLアイテム詳細_開発テスト用.xlsx] として提供しておりますのであわせてご参照ください。また、description 列の更新は、VBAマクロのサンプルを[UpdateDescription_SampleVBA.txt] として提供しておりますので、必要に応じてご参照ください。

次のファイルが含まれます。ファイル一式は[こちらから](https://github.com/EsriJapan/bulk_update_agol_items/archive/refs/heads/main.zip) 入手可能です。
- ノートブック ファイル： Update_Items_fromExcel.ipynb
- Excel のサンプル： 3D都市モデル_AGOLアイテム詳細_開発テスト用.xlsx
- description 列を更新するVBAマクロのサンプル： UpdateDescription_SampleVBA.txt

## 免責事項
* 本リポジトリに含まれるノートブック ファイルはサンプルとして提供しているものであり、動作に関する保証、および製品ライフサイクルに従った Esri 製品サポート サービスは提供しておりません。
* 本ツールに含まれるツールによって生じた損失及び損害等について、一切の責任を負いかねますのでご了承ください。
* 弊社で提供しているEsri 製品サポートサービスでは、本ツールに関しての Ｑ＆Ａ サポートの受付を行っておりませんので、予めご了承の上、ご利用ください。詳細は[
ESRIジャパン GitHub アカウントにおけるオープンソースへの貢献について](https://github.com/EsriJapan/contributing)をご参照ください。

## ライセンス
Copyright 2022 Esri Japan Corporation.

Apache License Version 2.0（「本ライセンス」）に基づいてライセンスされます。あなたがこのファイルを使用するためには、本ライセンスに従わなければなりません。
本ライセンスのコピーは下記の場所から入手できます。

> http://www.apache.org/licenses/LICENSE-2.0

適用される法律または書面での同意によって命じられない限り、本ライセンスに基づいて頒布されるソフトウェアは、明示黙示を問わず、いかなる保証も条件もなしに「現状のまま」頒布されます。本ライセンスでの権利と制限を規定した文言については、本ライセンスを参照してください。

ライセンスのコピーは本リポジトリの[ライセンス ファイル](./LICENSE)で利用可能です。