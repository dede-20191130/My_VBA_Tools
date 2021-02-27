VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSettingFormCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Class")
Option Compare Database
Option Explicit

'**************************
'*セッティングForm処理クラス　共通処理
'**************************

'定数欄
Private Const SOURCE_NAME As String = "clsSettingFormCommon"

'変数欄



'******************************************************************************************
'*getter/setter欄
'******************************************************************************************




'******************************************************************************************
'*関数名    ：deleteCurrentRcd
'*機能      ：
'*引数      ：
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function deleteCurrentRcd() As Boolean
    
    '定数
    Const FUNC_NAME As String = "deleteCurrentRcd"
    
    '変数
    Dim currentID As Long
    Dim wrs As New clsWrappedRecordSet
    
    On Error GoTo ErrorHandler

    deleteCurrentRcd = False
    
    'レコードの存在
    If (Form_F03_設定.subF_Datasheet.Form.recordset.BOF And Form_F03_設定.subF_Datasheet.Form.recordset.EOF) Or Form_F03_設定.subF_Datasheet.Form.recordset.RecordCount = 0 Then
        MsgBox "レコードが存在しません。", vbCritical, MESSAGE_TITLE_WARNING
        GoTo TruePoint
    End If
    
    currentID = Form_F03_設定.subF_Datasheet.Form.txt_ID.VALUE
    
    '確認
    If MsgBox("ID:" & currentID & "のレコードを消去します。" & vbNewLine & "よろしいですか。", vbYesNo, TOOL_NAME) <> vbYes Then GoTo ExitHandler
    
    '削除
    Set wrs.varRecordset = gObjDB.Database.OpenRecordset( _
        "SELECT * FROM" & _
        myVBVacant & _
        Form_F03_設定.gSubFormRcdSourceObj & _
        myVBVacant & _
        "WHERE" & _
        myVBVacant & _
        Form_F03_設定.gIDFieldName & _
        myVBVacant & _
        "=" & _
        myVBVacant & _
        currentID & _
        ";" _
        )
    With wrs.varRecordset
        .Delete
    End With
    
    
    Form_F03_設定.Refresh
 
TruePoint:

    deleteCurrentRcd = True
    
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

