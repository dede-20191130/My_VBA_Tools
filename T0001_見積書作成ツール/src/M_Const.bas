Attribute VB_Name = "M_Const"
'@Folder("VBAProject")
Option Explicit

'# tool name
Public Const TOOL_NAME As String = "T0001_見積書作成ツール"

'# シート名前
Public Const SHEET_NAME_ESTIMATE_DATA As String = "見積データ"
Public Const SHEET_NAME_ESTIMATE_PRODUCT_SET_DATA As String = "見積商品セットデータ"
Public Const SHEET_NAME_PRODUCT_DATA As String = "商品データ"
Public Const SHEET_NAME_BASIC_DATA As String = "基礎データ"
Public Const SHEET_NAME_TEMPLATE As String = "テンプレート"

'# シートオブジェクト
Public ws_estimate_data As Worksheet
Public ws_estimate_product_set_data As Worksheet
Public ws_product_data As Worksheet
Public ws_basic_data As Worksheet
Public ws_template As Worksheet

'OLEオブジェクト呼び出し名
Public Const OP_BUTTON_PDF_EXPORT_OFF As String = "OptionButton_PDF_Export_OFF"
Public Const OP_BUTTON_PDF_EXPORT_ON As String = "OptionButton_PDF_Export_ON"

'# 名前定義のキー
Public Const STR_NAME_RANGE_ESTIMATE_DATA_HEDD As String = "見積データ列_ヘッダー"
Public Const STR_NAME_RANGE_PRODUCT_DATA_HEDD As String = "商品データ列_ヘッダー"
Public Const STR_NAME_RANGE_ESTIMATE_PRODUCT_SET_DATA_HEDD As String = "見積商品セットデータ列_ヘッダー"
Public Const STR_NAME_RANGE_PRODUCT_TABLE_HEDD As String = "商品テーブル_ヘッダー"
Public Const STR_NAME_RANGE_ESTIMATE_DATA_SET_NUM_FIELD As String = "見積データ_見積商品セット番号入力欄"
Public Const STR_NAME_RANGE_ESTIMATE_PRODUCT_SET_DATA_CODE_FIELD As String = "見積商品セットデータ_商品コード入力欄"
Public Const STR_NAME_RANGE_CONSUME_TAX As String = "消費税の値"
Public Const STR_NAME_RANGE_TMPL_OTHER_NOTES As String = "テンプレート_その他備考格納セル"
Public Const STR_NAME_RANGE_TMPL_COMPANY As String = "テンプレート_会社名格納セル"
Public Const STR_NAME_RANGE_TMPL_ESTIMATE_DATE As String = "テンプレート_見積日格納セル"
Public Const STR_NAME_RANGE_TMPL_NUMBER As String = "テンプレート_見積番号格納セル"
Public Const STR_NAME_RANGE_TMPL_PAYMENT_TERMS As String = "テンプレート_支払条件格納セル"
Public Const STR_NAME_RANGE_TMPL_PERSON_IN_CHARGE As String = "テンプレート_担当者格納セル"
Public Const STR_NAME_RANGE_TMPL_DELIVERY_DATE As String = "テンプレート_納期格納セル"
Public Const STR_NAME_RANGE_TMPL_EXPIRATION_DATE As String = "テンプレート_有効期限格納セル"
Public Const STR_NAME_RANGE_TMPL_TAX As String = "テンプレート_消費税格納セル"

'# エラー文言
Public Const ITEM_KEY_FOR_ERR_MSG As String = "$1"

Public Const ERR_MSG_CREATED_DOCS_MAIN_HEDD As String = "下記の入力ミスが存在します。" & _
vbLf & _
"エラー内容に従って修正をしてください。"
Public Const ERR_MSG_CREATED_DOCS_MAIN_EACH_HEDD As String = "下記のデータ番号の見積書は正常に作成されませんでした。"
Public Const ERR_MSG_INVALID_VALUE As String = "【$1】" & vbTab & "有効な値が入力されていません。"
Public Const ERR_MSG_INCONSISTENCY_OF_DATA_NUM As String = "【$1】" & vbTab & "はじめの値は終わりの値以下にしてください。"
Public Const ERR_MSG_ESTIMATE_DATA_MUST_INPUT As String = "【$1】" & vbTab & "見積データ表の必須項目が空欄です。"
Public Const ERR_MSG_ESTIMATE_PRODUCT_SET_DATA_IS_NULL As String = "【$1】" & vbTab & "指定された見積商品セット番号に対応するデータが存在しません。"
Public Const ERR_MSG_INVALID_INDEX As String = "ツールのインデックスが変更されている可能性があります。" _
& vbLf _
& "ツールのメンテナンスが必要です。"
Public Const ERR_MSG_CHANGE_EVENT_DUPLICATION As String = "重複する値は入力できません。"

'# ActiveXオブジェクト生成定数
Public Const STR_ACTIVEX_OBJ_FILE_SYSTEM_OBJ As String = "Scripting.Filesystemobject"
Public Const STR_ACTIVEX_OBJ_DICTIONARY As String = "Scripting.Dictionary"

'再利用するクラスオブジェクト
Public obj_set_data As Cls_Set_Data
Public obj_product_data As Cls_Product_Data

'# 開発用定数
'イベント無効化
Public Const EVT_DISABLE_FLG As Boolean = False




