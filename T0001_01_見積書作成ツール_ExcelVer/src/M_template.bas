Attribute VB_Name = "M_template"
'******************************************************************************************
'*関数名    ：XXX
'*機能      ：
'*引数(1)   ：
'******************************************************************************************
Public Sub XXX()
    
    '定数
    Const FUNC_NAME As String = "XXX"
    
    '変数
    
    On Error GoTo ErrorHandler
    '---以下に処理を記述---

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生しましたのでマクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Sub

'******************************************************************************************
'*関数名    ：YYY
'*機能      ：
'*引数(1)   ：
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function YYY() As Boolean
    
    '定数
    Const FUNC_NAME As String = "YYY"
    
    '変数
    
    On Error GoTo ErrorHandler
    '戻り値初期値
    YYY = False
    
    '---以下に処理を記述---


    '戻り値設定
    YYY = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "エラーが発生しましたのでマクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Function

'******************************************************************************************
'*関数名    ：XXX2
'*機能      ：
'*引数(1)   ：
'******************************************************************************************
Public Sub XXX2()
    
    '定数
    Const FUNC_NAME As String = "XXX2"
    
    '変数
    
    On Error GoTo ErrorHandler
    
    '---以下に処理を記述---
    
ExitHandler:

    Exit Sub
    
ErrorHandler:
    
    If InStr(Err.Description, "発生場所：") <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, _
                  Err.Source, _
                  "エラー詳細：" & Err.Description & vbNewLine & _
                  "発生場所：" & FUNC_NAME & vbNewLine & _
                  "行番号：" & Erl & "（0は行番号設定無し）"
    End If
        
End Sub

'******************************************************************************************
'*関数名    ：YYY2
'*機能      ：
'*引数(1)   ：
'*戻り値    ：文字列
'******************************************************************************************
Public Function YYY2() As String
    
    '定数
    Const FUNC_NAME As String = "YYY2"
    
    '変数
    
    On Error GoTo ErrorHandler
    
    '戻り値初期値
    YYY2 = ""
    
    '---以下に処理を記述---


    '戻り値設定
'    YYY2 = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:
    
    If InStr(Err.Description, "発生場所：") <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, _
                  Err.Source, _
                  "エラー詳細：" & Err.Description & vbNewLine & _
                  "発生場所：" & FUNC_NAME & vbNewLine & _
                  "行番号：" & Erl & "（0は行番号設定無し）"
    End If
    
    
End Function


