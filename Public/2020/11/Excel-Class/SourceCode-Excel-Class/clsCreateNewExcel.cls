VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCreateNewExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("class")
Option Compare Database
Option Explicit

'**************************
'*Excelブック作成クラス
'**************************

'定数
Private currentSheetNum As Long

'変数
Private myXlApp As Object
Private myWorkBook As Object
Private dicWorkSheet As Dictionary 'ブックの全てのワークシートが格納されているDictionaryオブジェクト


'******************************************************************************************
'*getter/setter
'******************************************************************************************
Public Property Get xlApplication() As Object
    Set xlApplication = myXlApp
End Property


Public Property Get Workbook() As Object
    Set Workbook = myWorkBook
End Property


Public Property Get WorkSheets(ByVal num As Long) As Object
    If Not dicWorkSheet.Exists(num) Then Call MsgBox("シートがありません。", vbExclamation, TOOL_NAME): Set WorkSheets = Nothing: Exit Property
    Set WorkSheets = dicWorkSheet.Item(num)
End Property


'******************************************************************************************
'*関数名    ：Class_Initialize
'*機能      ：
'*引数      ：
'******************************************************************************************
Private Sub Class_Initialize()
    
    '定数
    Const FUNC_NAME As String = "Class_Initialize"
    
    '変数
    
    On Error GoTo ErrorHandler
    
    'シート数の設定
    currentSheetNum = 1
    
    'オブジェクトの設定
    Set myXlApp = CreateObject("Excel.Application")
    With myXlApp
        'バックグラウンドで生成・編集
        .Visible = False
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With
    Set myWorkBook = myXlApp.Workbooks.Add
    Set dicWorkSheet = New Dictionary
    dicWorkSheet.Add currentSheetNum, myWorkBook.WorkSheets(currentSheetNum)
    
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
'*引数      ：
'******************************************************************************************
Private Sub Class_Terminate()
    
    '定数
    Const FUNC_NAME As String = "Class_Terminate"
    
    '変数
    
    On Error GoTo ErrorHandler
    
    'ブックの保存 To:=デスクトップ
    With CreateObject("WScript.Shell")
        myWorkBook.SaveAs .SpecialFolders("Desktop") & "\" & "Test-Excel-Class-" & Format(Now, "yyyymmddhhnnss") & ".xlsx"
    End With
    
    '設定のリストア
    With myXlApp
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With
    
    'クローズ
    myWorkBook.Close
    myXlApp.Quit

ExitHandler:
    
    Set dicWorkSheet = Nothing
    Set myWorkBook = Nothing
    Set myXlApp = Nothing
    
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
'*関数名    ：addNewSheet
'*機能      ：ブックにシートを追加
'*引数      ：
'*戻り値    ：追加されたシートオブジェクト
'******************************************************************************************
Public Function addNewSheet() As Object
    
    '定数
    Const FUNC_NAME As String = "addNewSheet"
    
    '変数
    Dim ws As Object
    
    On Error GoTo ErrorHandler

    Set addNewSheet = Nothing
    
    currentSheetNum = currentSheetNum + 1
    '最後尾にシート追加
    Set ws = myWorkBook.WorkSheets.Add(After:=myWorkBook.WorkSheets(myWorkBook.WorkSheets.Count))
    dicWorkSheet.Add currentSheetNum, ws
    
    Set addNewSheet = ws
    
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

