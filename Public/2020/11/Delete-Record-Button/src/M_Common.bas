Attribute VB_Name = "M_Common"
'@Folder("Database")
Option Compare Database
Option Explicit

'**************************************
'*一般関数モジュール
'**************************************


'定数

'変数



'******************************************************************************************
'*関数名    ：createDummyString
'*機能      ：ダミーデータのための10文字のランダム文字列（a~z,A~Z,その他記号）を生成
'*引数      ：
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function createDummyString() As String
    
    '定数
    Const FUNC_NAME As String = "createDummyString"
    
    '変数
    Dim rtnVal As String
    Dim i As Long
    
    On Error GoTo ErrorHandler

    createDummyString = ""

    Call Randomize
    
    rtnVal = String(10, vbNullChar)
    
    For i = 1 To 10
        Mid(rtnVal, InStr(rtnVal, vbNullChar), 1) = Chr(65 + Int(Rnd * (122 - 65 + 1)))
    Next i
    
    createDummyString = rtnVal
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Function

