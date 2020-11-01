VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPolymo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Explicit


'******************************************************************************************
'*関数名    ：btn_execute_Click_Current
'*機能      ：現在時刻を表示する
'*引数(1)   ：
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function btn_execute_Click_Current() As Boolean
    
    '定数
    Const FUNC_NAME As String = "btn_execute_Click_Current"
    
    '変数
    
    On Error GoTo ErrorHandler

    btn_execute_Click_Current = False
    
    '現在時刻を表示する
    MsgBox "現在時刻：" & Now, , Tool_Name

    btn_execute_Click_Current = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, Tool_Name
        
    GoTo ExitHandler
        
End Function


'******************************************************************************************
'*関数名    ：btn_execute_Click_User
'*機能      ：ユーザ名を表示する
'*引数(1)   ：
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function btn_execute_Click_User() As Boolean
    
    '定数
    Const FUNC_NAME As String = "btn_execute_Click_User"
    
    '変数
    
    On Error GoTo ErrorHandler

    btn_execute_Click_User = False
    
    With CreateObject("WScript.Network")
        'ユーザ名を表示する
        MsgBox "ユーザ名：" & .UserName, , Tool_Name
    End With

    btn_execute_Click_User = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, Tool_Name
        
    GoTo ExitHandler
        
End Function






'******************************************************************************************
'*関数名    ：btn_execute_Click_Greeting
'*機能      ：あいさつを表示する
'*引数(1)   ：
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function btn_execute_Click_Greeting() As Boolean
    
    '定数
    Const FUNC_NAME As String = "btn_execute_Click_Greeting"
    
    '変数
    
    On Error GoTo ErrorHandler

    btn_execute_Click_Greeting = False
    
    MsgBox "こんにちは。", , Tool_Name
    
    btn_execute_Click_Greeting = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, Tool_Name
        
    GoTo ExitHandler
        
End Function




