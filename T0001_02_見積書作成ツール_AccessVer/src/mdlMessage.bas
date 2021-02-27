Attribute VB_Name = "mdlMessage"
'@Folder("Database")
Option Compare Database
Option Explicit

'**************************************
'*メッセージモジュール
'**************************************


'定数

'変数



'******************************************************************************************
'*関数名    ：showErrMessageUpdateRcd
'*機能      ：レコード登録・更新時エラーメッセージ
'*引数      ：エラープロパティ
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function showErrMessageUpdateRcd(ByVal errNum As Long, ByVal errDescription As String) As Boolean
    
    '定数
    Const FUNC_NAME As String = "showErrMessageUpdateRcd"
    
    '変数
    
    On Error GoTo ErrorHandler

    showErrMessageUpdateRcd = False
    
    If errNum = 3022 Then MsgBox "この値はすでに登録済みです。", vbCritical, TOOL_NAME: GoTo TruePoint
    
    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & errNum & vbNewLine & _
           errDescription, vbCritical, TOOL_NAME

TruePoint:

    showErrMessageUpdateRcd = True
    
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

