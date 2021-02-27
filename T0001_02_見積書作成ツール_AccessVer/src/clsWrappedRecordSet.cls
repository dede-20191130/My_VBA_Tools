VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWrappedRecordSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Class")
Option Compare Database
Option Explicit

'**************************
'*RecordSetオブジェクトラッパークラス
'**************************

'定数欄
Private Const SOURCE_NAME As String = "clsWrappedRecordSet"

'変数欄
Private myRecordset As recordset


'******************************************************************************************
'*getter/setter欄
'******************************************************************************************
Public Property Set varRecordset(ByVal RHS As recordset)
    Set myRecordset = RHS
End Property


Public Property Get varRecordset() As recordset
    Set varRecordset = myRecordset
End Property



'******************************************************************************************
'*機能      ：Class_Terminate
'*引数      ：
'******************************************************************************************
Private Sub Class_Terminate()

    '定数
    Const FUNC_NAME As String = "Class_Terminate"
    
    '変数
    
    On Error GoTo ErrorHandler
    
    'close
    myRecordset.Close
    

ExitHandler:
    
    Set myRecordset = Nothing
    
    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & err.Number & vbNewLine & _
           err.description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Sub



