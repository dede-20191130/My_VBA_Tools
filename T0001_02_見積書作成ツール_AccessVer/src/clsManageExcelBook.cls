VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsManageExcelBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Class")
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
Private dicWorkSheet As Object 'ブックの全てのワークシートが格納されているDictionaryオブジェクト

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
    currentSheetNum = 0
    
    'オブジェクトの設定
    Set myXlApp = CreateObject("Excel.Application")
    With myXlApp
        'バックグラウンドで生成・編集
        .Visible = False
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With
    Set dicWorkSheet = CreateObject(SCRIPTING_DICTIONARY)
    
ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & err.Number & vbNewLine & _
           err.description, vbCritical, TOOL_NAME
        
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
    
    'ブックの保存 To:=指定フォルダ
    With CreateObject("WScript.Shell")
        myWorkBook.SaveAs myWorkBook.FullName
    End With
    
    '設定のリストア
    With myXlApp
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With

ExitHandler:
    
    'クローズ
    myWorkBook.Close False
    myXlApp.Quit
    
    Set dicWorkSheet = Nothing
    Set myWorkBook = Nothing
    Set myXlApp = Nothing
    
    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & err.Number & vbNewLine & _
           err.description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Sub



'******************************************************************************************
'*機能      ：既存のブックを開きインスタンスに追加
'*引数      ：ファイルパス
'*引数      ：パスワード
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function addExistingBook(ByVal wbPath As String, Optional ByVal pw As String = "") As Object
    
    '定数
    Const FUNC_NAME As String = "addExistingBook"
    
    '変数
    
    On Error GoTo ErrorHandler

    Set addExistingBook = Nothing
    
    'ブックを開く
    Set myWorkBook = myXlApp.Workbooks.Open(fileName:=wbPath, Password:=pw, WriteResPassword:=pw)
    
    '既存のシートをdicに格納
    Dim ws As Object
    For Each ws In myWorkBook.WorkSheets
        currentSheetNum = currentSheetNum + 1
        dicWorkSheet.Add currentSheetNum, ws
    Next ws
    
    Set addExistingBook = myWorkBook

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
    Set ws = myWorkBook.WorkSheets.Add(after:=myWorkBook.WorkSheets(myWorkBook.WorkSheets.Count))
    dicWorkSheet.Add currentSheetNum, ws
    
    Set addNewSheet = ws
    
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



