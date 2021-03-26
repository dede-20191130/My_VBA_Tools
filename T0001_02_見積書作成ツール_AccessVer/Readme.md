
# T0001_02_見積書作成ツール_AccessVer





<!-- START doctoc -->
<!-- END doctoc -->






## このツールについて

以前、  
こちらの記事で紹介したように、  
Excelベースの見積書作成ツールを作成しました。  

[T0001_01_見積書作成ツール_ExcelVer](https://github.com/dede-20191130/My_VBA_Tools/tree/master/T0001_01_%E8%A6%8B%E7%A9%8D%E6%9B%B8%E4%BD%9C%E6%88%90%E3%83%84%E3%83%BC%E3%83%AB_ExcelVer)

今回、ほぼ同機能を持つツールを、  
Accessベースに移植しました。

理由としては、  
- 一対多のリレーションシップを持つデータを管理、抽出するのは  
Accessのほうが遥かに容易である。
- GUIの作成のための機能に関して、Accessがよりリッチであり、  
直感的にも使用しやすいため。
- Accessでのツール作成技術の向上のため。

です。

## 概要

各画面で、  
あらかじめ設定したマスタデータを選択し、  
テンプレート見積書に設定したデータを挿入します。  

Excelブック形式で  
見積書を出力します。

出力した見積書に使用したデータを保持し、  
再利用できます。


## 動作環境
- Windows
- 2016以上のOfficeソフトが動作する環境であれば可。  
2013以下のOfficeでも動作する可能性はあります。  
（そこまでの下位Ver.互換性に需要があるかどうかは不明なため、検証はしていません）。

## ツール外観（各画面紹介）

[Click To Show Slide Page](https://dede-20191130.github.io/learnerBlog/posts/2021/03/17/create-estm-accessver-tool-screen/#1)

## 機能紹介

[Click To Show Slide Page](https://dede-20191130.github.io/learnerBlog/posts/2021/03/24/create-estm-accessver-tool-faculty/#1)

## ご要望について

　もし書式の変更や追加の機能の作成等のご要望がございましたら

　こちらを御覧ください。



[My_VBA_Toolsのご説明](https://github.com/dede-20191130/My_VBA_Tools#%E4%BB%95%E4%BA%8B%E3%81%AE%E3%81%94%E4%BE%9D%E9%A0%BC)

