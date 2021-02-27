Attribute VB_Name = "mdlCommon"
'@Folder("Database")
Option Compare Database
Option Explicit

'**************************************
'*一般関数モジュール
'**************************************


'定数

'変数




'******************************************************************************************
'*関数名    ：initializeTool
'*機能      ：
'*引数(1)   ：
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function initializeTool() As Boolean
    
    '定数
    Const FUNC_NAME As String = "initializeTool"
    
    '変数
    
    On Error GoTo ErrorHandler

    initializeTool = False
    
    'グローバル変数の設定
    Set gObjDB = New clsDB
    Set gObjDtTrsfrManager = New clsDtTrsfrManager

    initializeTool = True
    
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




'******************************************************************************************
'*機能      ：Collectionのアイテムを配列に変換
'*引数      ：対象
'*戻り値    ：配列
'******************************************************************************************
Public Function CollectionToArray(myCol As Collection) As Variant
    
    '定数
    Const FUNC_NAME As String = "CollectionToArray"
    
    '変数
    Dim result  As Variant
    Dim cnt     As Long
    
    CollectionToArray = False
    
    ReDim result(0 To myCol.Count - 1)

    For cnt = 0 To myCol.Count - 1
        result(cnt) = myCol(cnt + 1)
    Next cnt

    CollectionToArray = result


ExitHandler:

    Exit Function
    
End Function

