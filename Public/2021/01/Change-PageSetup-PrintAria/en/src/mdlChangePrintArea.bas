Attribute VB_Name = "mdlChangePrintArea"
Option Explicit

'**************************
'*PageSetup_PrintArea Changing Test Module
'**************************

'Consts
Private Const SOURCE_NAME As String = "mdlChangePrintArea"


'Vars




'******************************************************************************************
'*Function :it's a function Before Modified
'*          extend PrintArea to one line below
'******************************************************************************************
Public Sub changePrintAreaBeforeModified()
    
    'Consts
    Const FUNC_NAME As String = "changePrintAreaBeforeModified"
    
    'Vars
    Dim prePrintAreaAddress As String
    
    On Error GoTo ErrorHandler
    
    With ThisWorkbook.Worksheets(1)
    
        'Current Print Area Address
        prePrintAreaAddress = .PageSetup.PrintArea
        
        'extend PrintArea to one line below
        .PageSetup.PrintArea = .Range(prePrintAreaAddress).Resize(.Range(prePrintAreaAddress).Rows.Count + 1).Address
        
        Debug.Print .PageSetup.PrintArea
        
    End With

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, SOURCE_NAME
        
    GoTo ExitHandler
        
End Sub


'******************************************************************************************
'*Function :it's a function after midification of pattern No.1
'*          extend PrintArea to one line below
'******************************************************************************************
Public Sub changePrintAreaModified01()
    
    'Consts
    Const FUNC_NAME As String = "changePrintAreaModified01"
    
    'Vars
    Dim prePrintAreaAddress As String
    Dim currentStyle As XlReferenceStyle

    On Error GoTo ErrorHandler
    
    With ThisWorkbook.Worksheets(1)
        
        'change the reference style to A1 style
        currentStyle = Application.ReferenceStyle
        Application.ReferenceStyle = xlA1
        
        'Current Print Area Address
        prePrintAreaAddress = .PageSetup.PrintArea
        
        'extend PrintArea to one line below
        .PageSetup.PrintArea = .Range(prePrintAreaAddress).Resize(.Range(prePrintAreaAddress).Rows.Count + 1).Address
        
        Debug.Print .PageSetup.PrintArea
        
        'restore the reference style
        Application.ReferenceStyle = currentStyle
        
    End With

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, SOURCE_NAME
        
    GoTo ExitHandler
        
End Sub


'******************************************************************************************
'*Function :it's a function after midification of pattern No.2
'*          extend PrintArea to one line below
'******************************************************************************************
Public Sub changePrintAreaModified02()
    
    'Consts
    Const FUNC_NAME As String = "changePrintAreaModified02"
    
    'Vars
    Dim prePrintAreaAddress As String
    
    On Error GoTo ErrorHandler
    
    With ThisWorkbook.Worksheets(1)
    
        'Current Print Area Address
        prePrintAreaAddress = .PageSetup.PrintArea
        
        'modify the address of prePrintAreaAddress to xlR1C1 style
        '** it doesn't change application's reference style
        If Application.ReferenceStyle = xlR1C1 Then prePrintAreaAddress = Application.ConvertFormula(prePrintAreaAddress, xlR1C1, xlA1)
        
        'extend PrintArea to one line below
        .PageSetup.PrintArea = .Range(prePrintAreaAddress).Resize(.Range(prePrintAreaAddress).Rows.Count + 1).Address
        
        Debug.Print .PageSetup.PrintArea
        
    End With

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, SOURCE_NAME
        
    GoTo ExitHandler
        
End Sub





