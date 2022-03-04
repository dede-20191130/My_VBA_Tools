Attribute VB_Name = "M_Test"
Option Explicit



'**************************
'*Exception Handling Sample Module
'**************************






'******************************************************************************************
'*Function :exception handling sample main
'******************************************************************************************
Public Sub main()
    
    'Consts
    Const FUNC_NAME As String = "main"
    
    'Vars
    Dim filePathArr As Variant
    Dim filePath As Variant
    Dim sheetName As String
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    'call funcSample01 with a sheet name which doesn't exist as a argument.
    sheetName = "sheetNotExist"
    filePathArr = funcSample01(sheetName)
    'show message if Null value is returned
    If IsNull(filePathArr) Then MsgBox sheetName & "The '" & sheetName & "' sheet doesn't exist." & vbNewLine & "Failed to retrieve the file path array, but the process continues."
    
    'call funcSample01 with a sheet name which exists as a argument.
    sheetName = "FilePath"
    filePathArr = funcSample01(sheetName)
    'show message if Null value is returned
    If IsNull(filePathArr) Then MsgBox sheetName & "The '" & sheetName & "' sheet doesn't exist." & vbNewLine & "Failed to retrieve the file path array, but the process continues."
    
    'call funcSample02 with each excel file path
    For Each filePath In filePathArr
        'if there is already some text in A1 cell, output the path in which the process failed to write into Immediate Window
        If Not funcSample02(ThisWorkbook.Path & filePath) Then
            Debug.Print "The file path in which the process failed to write: " & filePath
        End If
    Next filePath
    
    'the other errors not caught by funcSamples are caught by ErrorHandler labeded line in this procedure

ExitHandler:
    
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Macro"
        
    GoTo ExitHandler
        
End Sub




'******************************************************************************************
'*Function :example of function procedure containing a exception handling no1
'*          get a array of file paths
'*Arg      :worksheet name
'*Return   :array > normal termination; Null > abnormal termination
'******************************************************************************************
Public Function funcSample01(ByVal wsName As String) As Variant
    
    'Consts
    Const FUNC_NAME As String = "funcSample01"
    
    
    On Error GoTo ErrorHandler

    funcSample01 = Null
    
    'get a array of the values from A1 cell to A3 cell
    With ThisWorkbook.Worksheets(wsName)
        funcSample01 = .Range("A1:A3").Value
    End With

ExitHandler:
    
    
    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Macro"
        
    GoTo ExitHandler
        
End Function



'******************************************************************************************
'*Function :example of function procedure containing a exception handling no1
'*          open a excel file whose path is given as a argument
'*          write current time in A1 cell of first sheet
'*          if second sheet exists, write 'Completed' in A1 cell of it
'*Arg      :the excel file path
'*Return   :True > normal termination; False > abnormal termination
'******************************************************************************************
Public Function funcSample02(ByVal filePath As String) As Boolean
    
    'Consts
    Const FUNC_NAME As String = "funcSample02"
    
    'Vars
    Dim wb As Workbook
    
    On Error GoTo ErrorHandler

    funcSample02 = False
    
    Set wb = Workbooks.Open(filePath)
    
    
    With wb
        'write current time
        'a error occurs if there is already a text in A1. This is an abnormal termination
        If Trim(.Worksheets(1).Range("A1").Value) <> "" Then Err.Raise 1000, , "There is already a text in A1 Cell."
        .Worksheets(1).Range("A1").Value = Now
        
        'this process terminates normally if second sheet doesn't exist
        If .Worksheets.Count < 2 Then GoTo TruePoint
        
        'write 'Completed'
        .Worksheets(2).Range("A1").Value = "Completed"
        
    End With
    

TruePoint:
    
    'save the book
    wb.Save
    
    funcSample02 = True

ExitHandler:
    
    'never fail to close the book whether if this process terminates normally or abnormally.
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    
    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Macro"
        
    GoTo ExitHandler
        
End Function





