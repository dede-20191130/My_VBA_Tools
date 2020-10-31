VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormatExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Explicit

'**************************************
'*ExcelファイルのA1カーソル移動クラス
'**************************************


'定数

'変数
Private myTgtFileCnt As Long                     '処理済みのブック件数


'getter/setter
Public Property Get tgtFileCnt() As Long
    tgtFileCnt = myTgtFileCnt
End Property


#If Not CBool(DEBUG_MODE) Then
Public Property Let tgtFileCnt(ByVal num As Long)
    myTgtFileCnt = num
End Property


#End If


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
    
    'イベント一時停止
    '画面描画一時停止
    '警告を省略
    With Application
        .EnableEvents = False
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TOOL_NAME
        
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
    
    'Applicationプロパティの復旧
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Sub


'******************************************************************************************
'*関数名    ：btn_execute_Click_ON
'*機能      ：
'*引数(1)   ：フォルダパス
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function btn_execute_Click_ON(ByVal pFolderPath As String) As Boolean
    
    '定数
    Const FUNC_NAME As String = "btn_execute_Click_ON"
    
    '変数
    Dim fso As Object
    
    On Error GoTo ErrorHandler

    btn_execute_Click_ON = False
    
    'FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    '処理数の初期化
    myTgtFileCnt = 0
    
    '再帰的なExcelファイルの探索・処理
    If Not folderSearchRecurse(pFolderPath, fso) Then GoTo ExitHandler

    btn_execute_Click_ON = True
    
ExitHandler:
    
    Set fso = Nothing
    
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


'******************************************************************************************
'*関数名    ：btn_execute_Click_OFF
'*機能      ：
'*引数(1)   ：フォルダパス
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function btn_execute_Click_OFF(ByVal pFolderPath As String) As Boolean
    
    '定数
    Const FUNC_NAME As String = "btn_execute_Click_OFF"
    
    '変数
    Dim fso As Object

    
    On Error GoTo ErrorHandler

    btn_execute_Click_OFF = False
    
    'FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    '処理数の初期化
    myTgtFileCnt = 0
    
    'Excelファイルの探索・処理
    If Not folderSearch(pFolderPath, fso) Then GoTo ExitHandler

    btn_execute_Click_OFF = True
    
ExitHandler:
    
    Set fso = Nothing
    
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


'******************************************************************************************
'*関数名    ：folderSearchRecurse
'*機能      ：再帰的にフォルダを探索し、Excelファイルを処理する
'*引数(1)   ：フォルダパス
'*引数(2)   ：FileSystemObject
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function folderSearchRecurse( _
       ByVal pFolderPath As String, _
       ByVal fso As Object _
       ) As Boolean
    
    '定数
    Const FUNC_NAME As String = "folderSearchRecurse"
    
    '変数
    Dim folder As Object
    Dim subfolder As Object
    Dim file As Object
  
    On Error GoTo ErrorHandler

    folderSearchRecurse = False
    
    Set folder = fso.GetFolder(pFolderPath)
    
    'フォルダ内のサブフォルダを列挙
    '（サブフォルダがなければループ内は通らず）
    For Each subfolder In folder.SubFolders
        '再帰的呼び出し
        folderSearchRecurse subfolder.Path, fso
    Next subfolder
  
    'カレントフォルダ内のファイルを列挙
    
    For Each file In folder.Files
        With file
            'A1カーソル移動関数呼出
            If Not Me.formatExcel(file.Path, fso) Then GoTo ExitHandler
        End With
    Next file


    folderSearchRecurse = True
    
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




'******************************************************************************************
'*関数名    ：folderSearchRecurse
'*機能      ：再帰的にフォルダを探索し、Excelファイルを処理する
'*引数(1)   ：フォルダパス
'*引数(2)   ：FileSystemObject
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function folderSearch( _
       ByVal pFolderPath As String, _
       ByVal fso As Object _
       ) As Boolean
    
    '定数
    Const FUNC_NAME As String = "folderSearch"
    
    '変数
    Dim folder As Object
    Dim subfolder As Object
    Dim file As Object
  
    On Error GoTo ErrorHandler

    folderSearch = False
    
    Set folder = fso.GetFolder(pFolderPath)
  
    'カレントフォルダ内のファイルを列挙
    For Each file In folder.Files
        With file
            'A1カーソル移動関数呼出
            If Not Me.formatExcel(file.Path, fso) Then GoTo ExitHandler
        End With
    Next file


    folderSearch = True
    
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




'******************************************************************************************
'*関数名    ：formatExcel
'*機能      ：Excelブックの体裁を整える
'               すべてのシートのカーソルを左上セル（A1セル）に合わせ、一番左のシートをアクティブにした状態にする。
'*引数(1)   ：ファイルパス
'*引数(2)   ：FileSystemObject
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function formatExcel( _
       ByVal pFilePath As String, _
       ByVal fso As Object _
       ) As Boolean
    
    '定数
    Const FUNC_NAME As String = "formatExcel"
    
    '変数
    Dim extension As String
    Dim wb As Workbook
    Dim cntObj As Worksheet
    
    On Error GoTo ErrorHandler

    formatExcel = False
    
    '拡張子がExcelではない場合は終了
    extension = fso.GetExtensionName(pFilePath)
    If Not ( _
       extension = "xls" Or _
       extension = "xlsx" Or _
       extension = "xlsm" _
       ) Then formatExcel = True: GoTo ExitHandler
    
    '一時ファイルの場合は終了
    If Left(fso.GetFileName(pFilePath), 2) = "~$" Then formatExcel = True: GoTo ExitHandler
    
    'ブックを開く
    'エクセルとして開けない場合は終了
    On Error Resume Next
    Set wb = Application.Workbooks.Open(pFilePath)
    If Err.Number <> 0 Then Err.Clear: formatExcel = True: GoTo ExitHandler
    Err.Clear
    On Error GoTo ErrorHandler
    
    With wb
        'すべてのシートのカーソルを左上セル（A1セル）に合わせる
        For Each cntObj In .Worksheets
            '非表示セルは飛ばす
            If cntObj.Visible = True Then
                'A1を選択
                cntObj.Select
                cntObj.Range("A1").Select
                '倍率を100に変更
                ActiveWindow.Zoom = 100
            End If
        Next
    
        '可視シートのうち最も順列の若いものを選択
        For Each cntObj In .Worksheets
            If cntObj.Visible = True Then
                cntObj.Select
                Exit For
            End If
        Next
    End With
    
    'Save
    wb.Save
    
    'インクリメント
    myTgtFileCnt = myTgtFileCnt + 1
    
    formatExcel = True
    
ExitHandler:
    
    If Not wb Is Nothing Then wb.Close False: Set wb = Nothing
    
    Set cntObj = Nothing
    
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



