VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsF09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Class")
Option Compare Database
Option Explicit

'**************************
'*見積書_項目リストレコード追加Form処理クラス
'**************************

'定数欄
Private Const SOURCE_NAME As String = "clsF09"

'変数欄
Private arrObjCmbEvent(0 To 1) As clsCmbListener
Private objTxtListenerNumberBox As clsTxtListenerNumberBox

'******************************************************************************************
'*getter/setter欄
'******************************************************************************************




'******************************************************************************************
'*機能      ：Form_Load処理
'*引数      ：フォーム
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function formLoad(ByVal pForm As Access.Form) As Boolean
    
    '定数
    Const FUNC_NAME As String = "formLoad"
    
    '変数
    
    On Error GoTo ErrorHandler

    formLoad = False
    
    'イベントリスナー
    Set arrObjCmbEvent(0) = New clsCmbListener: Set arrObjCmbEvent(0).ComboBox = pForm.cmb_prod
    Set arrObjCmbEvent(1) = New clsCmbListener: Set arrObjCmbEvent(1).ComboBox = pForm.cmb_unit
    Set objTxtListenerNumberBox = New clsTxtListenerNumberBox: Set objTxtListenerNumberBox.TextBox = pForm.txt_unitNum


TruePoint:

    formLoad = True

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




'******************************************************************************************
'*機能      ：btn_showDetail_Click処理
'*引数      ：フォーム
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function btnShowDetailClick(ByVal pForm As Access.Form) As Boolean
    
    '定数
    Const FUNC_NAME As String = "btnShowDetailClick"
    
    '変数
    
    On Error GoTo ErrorHandler

    btnShowDetailClick = False
    
    '商品　選択済みであること
    If IsNull(pForm.cmb_prod.VALUE) Then MsgBox "商品が未選択です。", vbExclamation, MESSAGE_TITLE_WARNING: GoTo TruePoint
    
    '詳細表示の画面に遷移
    DoCmd.OpenForm FormName_10, , , , , acDialog, _
        pForm.Name & ";" & pForm.cmb_prod.VALUE

TruePoint:

    btnShowDetailClick = True

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





'******************************************************************************************
'*機能      ：btn_addNew_Click処理
'*引数      ：フォーム
'*引数      ：データベース
'*引数      ： W_見積_表作成用データ　次の採番番号
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function btnAddNewClick(ByVal pForm As Access.Form, ByVal daoDB As dao.Database, ByVal nextNumber As Long) As Boolean
    
    '定数
    Const FUNC_NAME As String = "btnAddNewClick"
    
    '変数
    Dim boolErr As Boolean
    Dim wrs As New clsWrappedRecordSet
    
    On Error GoTo ErrorHandler

    btnAddNewClick = False
    
    With pForm
        'チェック処理
        If Not checkWhetherControlsVacant( _
            boolErr, .cmb_prod.VALUE, .txt_unitNum.VALUE, .cmb_unit.VALUE _
        ) Then GoTo ExitHandler
        If boolErr Then Call MsgBox(MESSAGE_EXIST_BLANK, vbExclamation, MESSAGE_TITLE_WARNING): GoTo TruePoint
    End With
    
    '表作成用データの一時テーブルにデータ格納
    Set wrs.varRecordset = daoDB.OpenRecordset(TBL_W_ESTM)
    With wrs.varRecordset
        .AddNew
        
        .Fields("Number").VALUE = nextNumber
        .Fields("商品コード").VALUE = pForm.cmb_prod.VALUE
        .Fields("数量").VALUE = pForm.txt_unitNum.VALUE
        .Fields("数量単位").VALUE = pForm.cmb_unit.VALUE
        If Nz(pForm.txt_other.VALUE, "") <> "" Then .Fields("備考").VALUE = pForm.txt_other.VALUE
        
        .Update
    End With
    
    '画面遷移
    DoCmd.Close acForm, pForm.Name, acSaveNo

    
TruePoint:

    btnAddNewClick = True

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

