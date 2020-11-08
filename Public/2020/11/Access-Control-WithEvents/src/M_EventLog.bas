Attribute VB_Name = "M_EventLog"
'@Folder("Module")
Option Compare Database
Option Explicit


'**************************
'*イベントログModule
'**************************

'定数


'変数
Public targetTxtBox As Access.TextBox


'******************************************************************************************
'*関数名    ：writeEventLogs
'*機能      ：テキストボックスにイベントログを書き込む
'*引数(1)   ：記入文字列
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function writeEventLogs(ByVal logTxt As String) As Boolean
    
    '定数
    Const FUNC_NAME As String = "writeEventLogs"
    
    '変数
    
    On Error GoTo ErrorHandler

    writeEventLogs = False
    
    If targetTxtBox.Value <> "" Then targetTxtBox.Value = targetTxtBox.Value & vbNewLine
    targetTxtBox.Value = targetTxtBox.Value & _
                         Now & _
                         " : " & _
                         logTxt
    
    writeEventLogs = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Access-Control-WithEvents"
        
    GoTo ExitHandler
        
End Function



