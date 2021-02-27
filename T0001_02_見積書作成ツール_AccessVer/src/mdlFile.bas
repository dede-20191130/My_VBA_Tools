Attribute VB_Name = "mdlFile"
'@Folder("Database")
Option Compare Database
Option Explicit

'**************************
'*ファイル操作モジュール
'**************************

'定数欄
Private Const SOURCE_NAME As String = "mdlFile"

'変数欄



'******************************************************************************************
'*getter/setter欄
'******************************************************************************************




'******************************************************************************************
'*機能      ：見積書テンプレートを指定したファイルパスとして保存
'*引数      ：Database
'*引数      ：ファイルパス
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function saveBookEstmTmpl(ByVal daoDB As Database, ByVal filePath As String) As Boolean
    
    '定数
    Const FUNC_NAME As String = "saveBookEstmTmpl"
    
    '変数
    Dim wrs As New clsWrappedRecordSet
    
    On Error GoTo ErrorHandler

    saveBookEstmTmpl = False
    
    'テンプレートを取得
    Set wrs.varRecordset = daoDB.OpenRecordset( _
        "SELECT FILE FROM " & _
        TBL_M_FILE & myVBVacant & _
        "WHERE FILENAME = 'template';" _
    )
       
    '保存
    With wrs.varRecordset
        .MoveFirst
        .Fields("FILE").VALUE.Fields("FileData").SaveToFile filePath
    End With
    
    
TruePoint:

    saveBookEstmTmpl = True

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

