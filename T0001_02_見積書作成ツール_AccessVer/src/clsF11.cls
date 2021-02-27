VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsF11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Class")
Option Compare Database
Option Explicit

'**************************
'*見積書_項目リストレコード削除Form処理クラス
'**************************

'定数欄
Private Const SOURCE_NAME As String = "clsF11"

'変数欄
Private myObjCmbListener As clsCmbListener



'******************************************************************************************
'*getter/setter欄
'******************************************************************************************



'******************************************************************************************
'*機能      ：Form_Load処理
'*引数      ：
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function formLoad(ByVal pForm As Access.Form) As Boolean
    
    '定数
    Const FUNC_NAME As String = "formLoad"
    
    '変数
    
    On Error GoTo ErrorHandler

    formLoad = False
    
    'イベントリスナー
    Set myObjCmbListener = New clsCmbListener: Set myObjCmbListener.ComboBox = pForm.cmb_Number

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
'*機能      ：btn_delete_Click処理
'*引数      ：
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function btnDeleteClick(ByVal pForm As Access.Form, ByVal daoDB As dao.Database) As Boolean
    
    '定数
    Const FUNC_NAME As String = "btnDeleteClick"
    
    '変数
    Dim boolErr As Boolean
    Dim wrs As New clsWrappedRecordSet
    
    On Error GoTo ErrorHandler

    btnDeleteClick = False
    
    With pForm
        'チェック処理
        If Not checkWhetherControlsVacant( _
            boolErr, .cmb_Number.VALUE _
        ) Then GoTo ExitHandler
        If boolErr Then Call MsgBox(MESSAGE_EXIST_BLANK, vbExclamation, MESSAGE_TITLE_WARNING): GoTo TruePoint
    End With
    
    '表作成用データの一時テーブルのレコードを削除
    Call daoDB.execute("DELETE FROM " & TBL_W_ESTM & myVBVacant & "WHERE [Number] = " & pForm.cmb_Number.VALUE & ";")
    
    'Number列の振り直し
    Set wrs.varRecordset = daoDB.OpenRecordset(TBL_W_ESTM)
    With wrs.varRecordset
        If Not .EOF Then
            .MoveFirst
            Dim i As Long: i = 1
            Do Until .EOF
                .EDIT
                
                .Fields("Number").VALUE = i: i = i + 1
                
                .Update
                .MoveNext
            Loop
        End If
    End With
    
    '画面遷移
    DoCmd.Close acForm, pForm.Name, acSaveNo


TruePoint:

    btnDeleteClick = True

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







