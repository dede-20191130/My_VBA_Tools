Attribute VB_Name = "M_ExceptionHandleTemplate"
Option Explicit

'**************************
'*Exception Handling Function Template
'**************************




'******************************************************************************************
'*Function :template for sub-procedure
'******************************************************************************************
Public Sub subTemplate()
    
    'Consts
    Const FUNC_NAME As String = "subTemplate"
    
    'Vars
    
    On Error GoTo ErrorHandler

    '---write processing---
    

ExitHandler:
    
    '---write termination processing---
    
    Exit Sub
    
ErrorHandler:
    
    '---write processing for excetion---
    '   - show message
    '   - write the sysmte error infomation into a logfile
    '   - create a e-mail to notice the system error and send it
    
    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Macro"
        
    GoTo ExitHandler
        
End Sub




'******************************************************************************************
'*Function :template for function-procedure no1
'*Return   :True > normal termination; False > abnormal termination
'******************************************************************************************
Public Function functionTemplate01() As Boolean
    
    'Consts
    Const FUNC_NAME As String = "functionTemplate01"
    
    'Vars
    
    On Error GoTo ErrorHandler

    functionTemplate01 = False
    
    '---write processing---

TruePoint:
    
    '---write termination processing only when normal termination---
    
    functionTemplate01 = True

ExitHandler:
    
    '---write termination processing---
    
    Exit Function
    
ErrorHandler:

    '---write processing for excetion---
    '   - show message
    '   - write the sysmte error infomation into a logfile
    '   - create a e-mail to notice the system error and send it
    
    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Macro"
        
    GoTo ExitHandler
        
End Function






'******************************************************************************************
'*Function :template for function-procedure no2
'*Return   :any type except for Null > normal termination; Null > abnormal termination
'******************************************************************************************
Public Function functionTemplate02() As Variant
    
    'Consts
    Const FUNC_NAME As String = "functionTemplate02"
    
    'Vars
    
    On Error GoTo ErrorHandler

    functionTemplate02 = Null
    
    '---write processing---

ExitHandler:
    
    '---write termination processing---
    
    Exit Function
    
ErrorHandler:

    '---write processing for excetion---
    '   - show message
    '   - write the sysmte error infomation into a logfile
    '   - create a e-mail to notice the system error and send it
    
    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Macro"
        
    GoTo ExitHandler
        
End Function



