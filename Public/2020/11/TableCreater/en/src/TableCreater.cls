VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableCreater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Class")
Option Explicit

'**************************
'*TableCreater
'**************************


'Const
Private Const HEADER_COLOR = 15917529            'header cell color

'Vars
Private myRange As Range                         'range of target table
Private myColumnSubTotal As Long                 'column number of subtotal


'******************************************************************************************
'*getter/setter
'******************************************************************************************


Public Property Set Range(ByVal pRng As Range)
    Set myRange = pRng
End Property


Public Property Let ColumnSubTotal(ByVal num As Long)
    'prohibit being refered with the Range isn't set yet
    If myRange Is Nothing Then Err.Raise 1000, , "The range is not set."
    'error if argument number is out of range of 'range' object.
    If num < myRange.EntireColumn(1).Column Or myRange.EntireColumn(myRange.EntireColumn.Count).Column < num Then Err.Raise 1001, , "Invalid column number specification."
    'set
    myColumnSubTotal = num
End Property


'******************************************************************************************
'*Function ：
'*Arg      ：
'******************************************************************************************
Private Sub Class_Initialize()
    
    'Const
    Const FUNC_NAME As String = "Class_Initialize"
    
    'Vars
    
    On Error GoTo ErrorHandler
    
    'There's nothing special to do here this Class.

ExitHandler:

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
'*Function ：draw lines
'*Arg      ：
'*Return   ：True > normal termination; False > abnormal termination
'******************************************************************************************
Public Function drawLines() As Boolean
    
    'Const
    Const FUNC_NAME As String = "drawLines"
    
    
    On Error GoTo ErrorHandler

    drawLines = False
    
    'prohibit being called with the Range isn't set yet
    If myRange Is Nothing Then Err.Raise 1000, , "The range is not set."
    
    'draw lines
    myRange.Borders.LineStyle = xlContinuous

    drawLines = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "TableCreater"
        
    GoTo ExitHandler
        
End Function


'******************************************************************************************
'*Function ：set header part's style
'            The header is cells of the first row of given range.
'*Return   ：True > normal termination; False > abnormal termination
'******************************************************************************************
Public Function setStyleForHeader() As Boolean
    
    'Const
    Const FUNC_NAME As String = "setStyleForHeader"
    
    
    On Error GoTo ErrorHandler

    setStyleForHeader = False
    
    'prohibit being called with the Range isn't set yet
    If myRange Is Nothing Then Err.Raise 1000, , "The range is not set."
    
    'change styles
    With myRange.Rows(1)
        'change background color
        .Interior.color = HEADER_COLOR
        'change font weight to bold
        .Font.Bold = True
        'change text alignment to center
        .HorizontalAlignment = xlCenter
    End With
        
    
    setStyleForHeader = True
    
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "TableCreater"
        
    GoTo ExitHandler
        
End Function




'******************************************************************************************
'*Function ：calculate total value from subtotal column and output it
'*Return   ：True > normal termination; False > abnormal termination
'******************************************************************************************
Public Function calcTotalFromSubTotal() As Boolean
    
    'Const
    Const FUNC_NAME As String = "calcTotalFromSubTotal"
    
    'Vars
    Dim sumVal As Long
    Dim cell As Range
    Dim subTotalOrder As Long
    
    On Error GoTo ErrorHandler

    calcTotalFromSubTotal = False
    
    'prohibit being called with the Range isn't set yet
    If myRange Is Nothing Then Err.Raise 1000, , "The range is not set."
    
    'prohibit being called with the column number for subtotal isn't set yet
    If myColumnSubTotal = 0 Then Err.Raise 1002, , "The column number for subtotal is not set."
        
    'calculate the order of subtotal column
    subTotalOrder = myColumnSubTotal - myRange(1).Column + 1
        
    'calculate total value, except for header row
    For Each cell In myRange.Columns(subTotalOrder).Cells.Offset(1).Resize(myRange.Columns(subTotalOrder).Cells.Offset(1).Cells.Count - 1)
        'add only numeric value
        If IsNumeric(cell.Value) Then sumVal = sumVal + cell.Value
    Next cell
    If sumVal = 0 Then GoTo TruePoint
    
    'write the total value in the bottom cell of subtotal column
    With myRange.Columns(subTotalOrder).Rows(myRange.Columns(subTotalOrder).Cells.Count).Offset(1)
        .Value = sumVal
        'refer the label cell
        With .Offset(, -1)
            'write a label
            .Value = "Total"
            'draw lines to label cell and total cell
            .Resize(, .Columns.Count + 1).Borders.LineStyle = xlContinuous
        End With
        
    End With
       
TruePoint:
       
    calcTotalFromSubTotal = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "TableCreater"
        
    GoTo ExitHandler
        
End Function


