Attribute VB_Name = "M_EventLog"
'@Folder("Module")
Option Compare Database
Option Explicit


'**************************
'*Event Log Module
'**************************

'Const


'Variable
Public targetTxtBox As Access.TextBox


'******************************************************************************************
'*Function Fwrite the event log into the textbox specified in a module variable
'*Arg(1)   Fthe written string
'*Return   FTrue > normal termination; False > abnormal termination

'******************************************************************************************
Public Function writeEventLogs(ByVal logTxt As String) As Boolean
    
    'Const
    Const FUNC_NAME As String = "writeEventLogs"
    
    'Variable
    
    On Error GoTo ErrorHandler

    writeEventLogs = False
    
    If Nz(targetTxtBox.Value, "") <> "" Then targetTxtBox.Value = targetTxtBox.Value & vbNewLine
    targetTxtBox.Value = targetTxtBox.Value & _
                         Now & _
                         " : " & _
                         logTxt
    
    writeEventLogs = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Access-Control-WithEvents"
        
    GoTo ExitHandler
        
End Function



