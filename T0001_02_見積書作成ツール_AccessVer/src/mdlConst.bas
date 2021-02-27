Attribute VB_Name = "mdlConst"
'@Folder("Database")
Option Compare Database
Option Explicit

'**************************************
'*Constモジュール
'**************************************


'定数

'変数


'*****ツールプロパティ
Public Const TOOL_NAME As String = "T0001_02_見積書作成ツール"

'*****VBのカスタマイズ定数
Public Const myVBVacant As String = " "
Public Const myVBUL As String = "_"
Public Const myVBSglQte As String = "'"


'*****ACTIVEX
Public Const SCRIPTING_DICTIONARY As String = "scripting.dictionary"


'*****テーブル
Public Const TBL_M_BSC_TAX As String = "M_基礎データ_消費税率"
Public Const TBL_M_BSC_UNIT As String = "M_基礎データ_数量単位"
Public Const TBL_M_MEMBER As String = "M_基礎データ_人名データ"
Public Const TBL_M_PROD As String = "M_基礎データ_商品データ"
Public Const TBL_M_ORG As String = "M_基礎データ_取引先会社データ"
Public Const TBL_M_FILE As String = "M_ファイル"
Public Const TBL_T_ESTM As String = "T_見積_表作成用データ"
Public Const TBL_T_ESTM_DTL As String = "T_見積書項目データ"
Public Const TBL_W_ESTM As String = "W_見積_表作成用データ"

'*****クエリ
Public Const QRY_QSR01 As String = "QRS01_人名"
Public Const QRY_Q_MAX_W_ESTM As String = "Q01_MAX_W_見積_表作成用データ"


'*****フォーム
Public Const FormName_01 As String = "F01_Init"
Public Const FormName_02 As String = "F02_メニュー"
Public Const FormName_03 As String = "F03_設定"
Public Const FormName_03_SUB01 As String = "F03_設定_SUB01_基礎データ_消費税率"
Public Const FormName_03_SUB02 As String = "F03_設定_SUB02_基礎データ_数量単位"
Public Const FormName_03_SUB03 As String = "F03_設定_SUB03_基礎データ_取引先会社データ"
Public Const FormName_03_SUB04 As String = "F03_設定_SUB04_基礎データ_人名データ"
Public Const FormName_03_SUB05 As String = "F03_設定_SUB05_基礎データ_商品データ"
Public Const FormName_04 As String = "F04_登録_編集"
Public Const FormName_04_SUB01 As String = "F04_登録_編集_SUB01_基礎データ_消費税率"
Public Const FormName_04_SUB02 As String = "F04_登録_編集_SUB02_基礎データ_数量単位"
Public Const FormName_04_SUB03 As String = "F04_登録_編集_SUB03_基礎データ_取引先会社データ"
Public Const FormName_04_SUB04 As String = "F04_登録_編集_SUB04_基礎データ_人名データ"
Public Const FormName_04_SUB05 As String = "F04_登録_編集_SUB05_基礎データ_商品データ"
Public Const FormName_05 As String = "F05_見積書項目設定"
Public Const FormName_06 As String = "F06_レコード選択"
Public Const FormName_07 As String = "F07_レコード選択_見積書項目"
Public Const FormName_08 As String = "F08_見積書_表作成"
Public Const FormName_09 As String = "F09_見積書_項目リストレコード追加"
Public Const FormName_10 As String = "F10_データシート表示"
Public Const FormName_11 As String = "F11_見積書_項目リストレコード削除"

'*****名前定義のキー
Public Const STR_NAME_RANGE_PRODUCT_TABLE_HEDD As String = "商品テーブル_ヘッダー"
Public Const STR_NAME_RANGE_TMPL_OTHER_NOTES As String = "テンプレート_その他備考格納セル"
Public Const STR_NAME_RANGE_TMPL_COMPANY As String = "テンプレート_会社名格納セル"
Public Const STR_NAME_RANGE_TMPL_ESTIMATE_DATE As String = "テンプレート_見積日格納セル"
Public Const STR_NAME_RANGE_TMPL_NUMBER As String = "テンプレート_見積番号格納セル"
Public Const STR_NAME_RANGE_TMPL_PAYMENT_TERMS As String = "テンプレート_支払条件格納セル"
Public Const STR_NAME_RANGE_TMPL_PERSON_IN_CHARGE As String = "テンプレート_担当者格納セル"
Public Const STR_NAME_RANGE_TMPL_DELIVERY_DATE As String = "テンプレート_納期格納セル"
Public Const STR_NAME_RANGE_TMPL_EXPIRATION_DATE As String = "テンプレート_有効期限格納セル"
Public Const STR_NAME_RANGE_TMPL_TAX As String = "テンプレート_消費税格納セル"

'*****メッセージ
Public Const MESSAGE_TITLE_NOTICE As String = "注意"
Public Const MESSAGE_TITLE_WARNING As String = "警告"
Public Const MESSAGE_TITLE_ERROR As String = "エラー"

'警告
Public Const MESSAGE_EXIST_BLANK As String = "入力欄に空欄が存在します。"

'*****エラー番号・メッセージ
Public Enum eNumCustomErr
    wrongArgs = 2000
End Enum

Public Const MSG_ERR_WRONG_ARGS As String = "引数が不正です。"


'*****固定文言
Public Const REGISTER As String = "登録"
Public Const EDIT As String = "編集"
