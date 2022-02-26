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
'*Excel Book Creation Class
'**************************

'Consts
Private currentSheetNum As Long

'Vars
Private myXlApp As Object
Private myWorkBook As Object
Private dicWorkSheet As Dictionary 'store all sheet objects of the book


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
    If Not dicWorkSheet.Exists(num) Then Call MsgBox("The Sheet does not exists.", vbExclamation, TOOL_NAME): Set WorkSheets = Nothing: Exit Property
    Set WorkSheets = dicWorkSheet.Item(num)
End Property


'******************************************************************************************
'******************************************************************************************
Private Sub Class_Initialize()
    
    'Consts
    Const FUNC_NAME As String = "Class_Initialize"
    
    'Vars
    
    On Error GoTo ErrorHandler
    
    'initial sheet number
    currentSheetNum = 1
    
    'instance of ExcelApp
    Set myXlApp = CreateObject("Excel.Application")
    With myXlApp
        'all processing are done n the background
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

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Sub


'******************************************************************************************
'******************************************************************************************
Private Sub Class_Terminate()
    
    'Consts
    Const FUNC_NAME As String = "Class_Terminate"
    
    'Vars
    
    On Error GoTo ErrorHandler
    
    'save the book to user's Desktop
    With CreateObject("WScript.Shell")
        myWorkBook.SaveAs .SpecialFolders("Desktop") & "\" & "Test-Excel-Class-" & Format(Now, "yyyymmddhhnnss") & ".xlsx"
    End With
    
    'restore ExcelApp settings
    With myXlApp
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With
    
    'close
    myWorkBook.Close
    myXlApp.Quit

ExitHandler:
    
    Set dicWorkSheet = Nothing
    Set myWorkBook = Nothing
    Set myXlApp = Nothing
    
    Exit Sub
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Sub



'******************************************************************************************
'*Function :add new sheet
'*Return   :added sheet
'******************************************************************************************
Public Function addNewSheet() As Object
    
    'Consts
    Const FUNC_NAME As String = "addNewSheet"
    
    'Vars
    Dim ws As Object
    
    On Error GoTo ErrorHandler

    Set addNewSheet = Nothing
    
    currentSheetNum = currentSheetNum + 1
    'add new sheet at the end
    Set ws = myWorkBook.WorkSheets.Add(After:=myWorkBook.WorkSheets(myWorkBook.WorkSheets.Count))
    dicWorkSheet.Add currentSheetNum, ws
    
    Set addNewSheet = ws
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Function

