Attribute VB_Name = "M_DBManager"
'@Folder("Database")
Option Compare Database
Option Explicit

'**************************
'*DB管理Module
'**************************

'定数


'変数


'******************************************************************************************
'*関数名    ：getAccessDB
'*機能      ：
'*引数(1)   ：
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function getAccessDB(ByVal dbFilePath As String) As DAO.Database
    
    '定数
    Const FUNC_NAME As String = "getAccessDB"
    
    '変数
    Dim pwStr As String
    Dim errFlg As Boolean
    
    On Error Resume Next
    Set getAccessDB = Nothing
    
    'データベースを開く
    Set getAccessDB = DBEngine.Workspaces(0).OpenDatabase(dbFilePath, , True)
    
    '正常完了時
    If Err.Number = 0 Then GoTo ExitHandler
    
    'パスワードが掛かっている場合
    If Err.Number = 3031 Then
        'エラーリセット
        Err.Clear
        'パスワードを入力させる
        pwStr = InputBox("Accessデータベースのパスワードを入力してください。", "パスワード入力")
        '再度データベースを開く
        Set getAccessDB = DBEngine.Workspaces(0).OpenDatabase(dbFilePath, False, True, "MS Access;PWD=" & pwStr)
        If Err.Number <> 0 Then errFlg = True
    Else
        errFlg = True
    End If
    
    'エラー発生時
    If errFlg Then GoTo ErrorHandler
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.description, vbCritical, Tool_Name
        
    GoTo ExitHandler
        
End Function



