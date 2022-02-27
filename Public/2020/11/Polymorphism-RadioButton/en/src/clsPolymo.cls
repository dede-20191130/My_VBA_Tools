VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPolymo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Explicit


'******************************************************************************************
'*Function :show current time
'*Arg(1)   :
'*Return   :True > normal termination; False > abnormal termination
'******************************************************************************************
Public Function btn_execute_Click_Current() As Boolean
    
    'Consts
    Const FUNC_NAME As String = "btn_execute_Click_Current"
    
    'Vars
    
    On Error GoTo ErrorHandler

    btn_execute_Click_Current = False
    
    'show current time
    MsgBox "Current time: " & Now, , Tool_Name

    btn_execute_Click_Current = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, Tool_Name
        
    GoTo ExitHandler
        
End Function


'******************************************************************************************
'*Function :show PC user name
'*Arg(1)   :
'*Return   :True > normal termination; False > abnormal termination
'******************************************************************************************
Public Function btn_execute_Click_User() As Boolean
    
    'Consts
    Const FUNC_NAME As String = "btn_execute_Click_User"
    
    'Vars
    
    On Error GoTo ErrorHandler

    btn_execute_Click_User = False
    
    With CreateObject("WScript.Network")
        'show PC user name
        MsgBox "Use name: " & .UserName, , Tool_Name
    End With

    btn_execute_Click_User = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, Tool_Name
        
    GoTo ExitHandler
        
End Function






'******************************************************************************************
'*Function :show greeting
'*Arg(1)   :
'*Return   :True > normal termination; False > abnormal termination
'******************************************************************************************
Public Function btn_execute_Click_Greeting() As Boolean
    
    'Consts
    Const FUNC_NAME As String = "btn_execute_Click_Greeting"
    
    'Vars
    
    On Error GoTo ErrorHandler

    btn_execute_Click_Greeting = False
    
    MsgBox "Hello.", , Tool_Name
    
    btn_execute_Click_Greeting = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, Tool_Name
        
    GoTo ExitHandler
        
End Function




