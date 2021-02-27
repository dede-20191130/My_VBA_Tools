VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Class")
Option Compare Database
Option Explicit

'**************************
'*自身のDB管理クラス
'**************************

'定数

'変数
Private myWorkspace As dao.Workspace
Private myDatabase As dao.Database
Private isTransactionBegin As Boolean

'******************************************************************************************
'*getter/setter
'******************************************************************************************
Public Property Get Database() As dao.Database
    Set Database = myDatabase
End Property


'******************************************************************************************
'*関数名    ：Class_Initialize
'*機能      ：
'*引数(1)   ：
'******************************************************************************************
Private Sub Class_Initialize()
    
    '定数
    Const FUNC_NAME As String = "Class_Initialize"
    
    '変数
    
    On Error GoTo ErrorHandler
    
    Set myWorkspace = Application.DBEngine.Workspaces(0)
    Set myDatabase = myWorkspace.Databases(0)
    isTransactionBegin = False
    
ExitHandler:

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




'******************************************************************************************
'*関数名    ：Class_Terminate
'*機能      ：
'*引数(1)   ：
'******************************************************************************************
Private Sub Class_Terminate()
    
    '定数
    Const FUNC_NAME As String = "Class_Terminate"
    
    '変数
    
    On Error GoTo ErrorHandler
    
    myDatabase.Close: Set myDatabase = Nothing
    myWorkspace.Close: Set myWorkspace = Nothing

ExitHandler:

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


'******************************************************************************************
'*トランザクション
'******************************************************************************************
Public Sub beginTransaction()
    If isTransactionBegin Then Exit Sub
    myWorkspace.BeginTrans
    isTransactionBegin = True
End Sub

Public Sub CommitTransaction()
    If Not isTransactionBegin Then Exit Sub
    myWorkspace.CommitTrans
    isTransactionBegin = False
End Sub

Public Sub RollbackTransaction()
    If Not isTransactionBegin Then Exit Sub
    myWorkspace.Rollback
    isTransactionBegin = False
End Sub
