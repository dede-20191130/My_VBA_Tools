Attribute VB_Name = "M_ExceptionHandleTemplate"
Option Explicit

'**************************
'*例外処理サンプル
'**************************




'******************************************************************************************
'*関数名    ：Subプロシージャの例外処理テンプレート
'*機能      ：
'*引数      ：
'******************************************************************************************
Public Sub subTemplate()
    
    '定数
    Const FUNC_NAME As String = "subTemplate"
    
    '変数
    
    On Error GoTo ErrorHandler

    '---ここから処理を記載する---
    

ExitHandler:
    
    '---ここから終了処理を記載する---
    
    Exit Sub
    
ErrorHandler:
    
    '---ここから例外発生時処理を記載する---
    '　　例：メッセージボックス表示、
    '　　　　ログファイルにシステムエラー情報書き込み、
    '　　　　システムエラー発生の通知メールの作成・発信など
    
    MsgBox "システムエラーが発生しました。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "マクロ"
        
    GoTo ExitHandler
        
End Sub




'******************************************************************************************
'*関数名    ：Functionプロシージャの例外処理テンプレート(1)
'*機能      ：
'*引数      ：
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function functionTemplate01() As Boolean
    
    '定数
    Const FUNC_NAME As String = "functionTemplate01"
    
    '変数
    
    On Error GoTo ErrorHandler

    functionTemplate01 = False
    
    '---ここから処理を記載する---

TruePoint:
    
    '---ここから正常時のみの終了処理を記載する---
    
    functionTemplate01 = True

ExitHandler:
    
    '---ここから終了処理を記載する---
    
    Exit Function
    
ErrorHandler:

    '---ここから例外発生時処理を記載する---
    '　　例：メッセージボックス表示、
    '　　　　ログファイルにシステムエラー情報書き込み、
    '　　　　システムエラー発生の通知メールの作成・発信など
    
    MsgBox "システムエラーが発生しました。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "マクロ"
        
    GoTo ExitHandler
        
End Function






'******************************************************************************************
'*関数名    ：Functionプロシージャの例外処理テンプレート(2)
'*機能      ：
'*引数      ：
'*戻り値    ：任意の指定の基本型 > 正常終了、Null > 異常終了
'******************************************************************************************
Public Function functionTemplate02() As Variant
    
    '定数
    Const FUNC_NAME As String = "functionTemplate02"
    
    '変数
    
    On Error GoTo ErrorHandler

    functionTemplate02 = Null
    
    '---ここから処理を記載する---

ExitHandler:
    
    '---ここから終了処理を記載する---
    
    Exit Function
    
ErrorHandler:

    '---ここから例外発生時処理を記載する---
    '　　例：メッセージボックス表示、
    '　　　　ログファイルにシステムエラー情報書き込み、
    '　　　　システムエラー発生の通知メールの作成・発信など
    
    MsgBox "システムエラーが発生しました。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "マクロ"
        
    GoTo ExitHandler
        
End Function

