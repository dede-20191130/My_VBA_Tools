VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSettingFormShoItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Class")
Option Compare Database
Option Explicit
Implements clsAbsSettingForm
'**************************
'*セッティングForm処理クラス　商品データ
'**************************

'定数欄
Private Const SOURCE_NAME As String = "clsSettingFormShoItem"

'変数欄
Private myObjSettingFormCommon As New clsSettingFormCommon


'******************************************************************************************
'*getter/setter欄
'******************************************************************************************
Private Property Get clsAbsSettingForm_objSettingFormCommon() As clsSettingFormCommon
    Set clsAbsSettingForm_objSettingFormCommon = myObjSettingFormCommon
End Property

Private Property Get clsAbsSettingForm_typeSettingForm() As eTypeSettingForm
    clsAbsSettingForm_typeSettingForm = eTypeSettingForm.shoItem
End Property


'******************************************************************************************
'*関数名    ：getInfo_numUnit
'*機能      ：フォームの情報を取得する　取引先会社データ
'*引数(1)   ：情報格納Dic
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Private Function clsAbsSettingForm_getInfo(pDicInfo As Object) As Boolean

    '定数
    Const FUNC_NAME As String = "clsAbsSettingForm_getInfo"
    
    '変数
    
    On Error GoTo ErrorHandler

    clsAbsSettingForm_getInfo = False
    
    'タイトル
    pDicInfo.Add "formTitle", "商品設定"
    'サブフォーム
    pDicInfo.Add "subFormSourceObject", FormName_03_SUB05
    'cmb_selectedRcdのRowSource
    pDicInfo.Add "rowSourceOfCmbSelectedRcd", "SELECT 商品ID FROM" & myVBVacant & TBL_M_PROD & myVBVacant & "ORDER BY 商品ID;"
    'フォームデータ構造体
    pDicInfo.Add "IDFirldName", "商品ID"
    pDicInfo.Add "subFormRcdSourceObj", TBL_M_PROD

    clsAbsSettingForm_getInfo = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & err.Number & vbNewLine & _
           err.description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Function





