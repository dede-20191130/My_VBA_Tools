Attribute VB_Name = "M_Caller"
'@Folder("Module")
Option Explicit

'**************************
'*calls TableCreater
'**************************

'Const
Const BASE_SHEET = "Base"

'Vars



'******************************************************************************************
'*getter/setter
'******************************************************************************************


'******************************************************************************************
'*Function ÅFcreate a table for template A in base sheet through TableCreater
'            creation location: new sheet
'******************************************************************************************
Public Sub TestTemplateA()
    
    'Const
    Const FUNC_NAME As String = "TestTemplateA"
    
    'Vars
    Dim ws As Worksheet
    Dim tableRange As Range
    Dim objTableCreater As TableCreater
    
    On Error GoTo ErrorHandler
    
    With ThisWorkbook
        'create new sheet
        Set ws = .Worksheets.Add(, .Worksheets(.Worksheets.Count))
        ws.Name = FUNC_NAME & "_" & Format(Now, "yyyymmddhhnnss")
        
        'copy template range from
        Set tableRange = ws.Range(ws.Cells(2, 2), ws.Cells(9, 4))
        tableRange.Value = .Worksheets(BASE_SHEET).Range(.Worksheets(BASE_SHEET).Cells(3, 2), .Worksheets(BASE_SHEET).Cells(10, 4)).Value
        
        'instanciate TableCreater
        Set objTableCreater = New TableCreater
        
        'set params
        Set objTableCreater.Range = tableRange
        objTableCreater.ColumnSubTotal = 4
        
        'draw lines: if error, shift to the exit process
        If Not objTableCreater.drawLines Then GoTo ExitHandler
         
        'set styles for header part for emphasis: if error, shift to the exit process
        If Not objTableCreater.setStyleForHeader Then GoTo ExitHandler
        
        'calc total: if error, shift to the exit process
        If Not objTableCreater.calcTotalFromSubTotal Then GoTo ExitHandler
        
        'adjust column widths
        tableRange.EntireColumn.AutoFit
        
    End With
    

ExitHandler:
    
    'release memory
    Set objTableCreater = Nothing
    Set ws = Nothing
    Set tableRange = Nothing
    
    Exit Sub
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "TableCreater"
        
    GoTo ExitHandler
        
End Sub






'******************************************************************************************
'*Function ÅFcreate a table for template B in base sheet through TableCreater
'            creation location: new sheet
'******************************************************************************************
Public Sub TestTemplateB()
    
    'Const
    Const FUNC_NAME As String = "TestTemplateB"
    
    'Vars
    Dim ws As Worksheet
    Dim tableRange As Range
    Dim objTableCreater As TableCreater
    
    On Error GoTo ErrorHandler
    
    With ThisWorkbook
        'create new sheet
        Set ws = .Worksheets.Add(, .Worksheets(.Worksheets.Count))
        ws.Name = FUNC_NAME & "_" & Format(Now, "yyyymmddhhnnss")
        
        'copy template range from
        Set tableRange = ws.Range(ws.Cells(2, 2), ws.Cells(8, 8))
        tableRange.Value = .Worksheets(BASE_SHEET).Range(.Worksheets(BASE_SHEET).Cells(13, 2), .Worksheets(BASE_SHEET).Cells(19, 8)).Value
        
        'instanciate TableCreater
        Set objTableCreater = New TableCreater
        
        'set params
        Set objTableCreater.Range = tableRange
        objTableCreater.ColumnSubTotal = 8
        
        'draw lines: if error, shift to the exit process
        If Not objTableCreater.drawLines Then GoTo ExitHandler
         
        'set styles for header part for emphasis: if error, shift to the exit process
        If Not objTableCreater.setStyleForHeader Then GoTo ExitHandler
        
        'calc total: if error, shift to the exit process
        If Not objTableCreater.calcTotalFromSubTotal Then GoTo ExitHandler
        
        'adjust column widths
        tableRange.EntireColumn.AutoFit
        
    End With
    

ExitHandler:
    
    'release memory
    Set objTableCreater = Nothing
    Set ws = Nothing
    Set tableRange = Nothing
    
    Exit Sub
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "TableCreater"
        
    GoTo ExitHandler
        
End Sub

