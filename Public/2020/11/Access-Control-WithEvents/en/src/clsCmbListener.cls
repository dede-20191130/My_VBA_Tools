VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCmbListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Class")
Option Compare Database
Option Explicit

'**************************
'*Combobox Event Listener
'**************************

'Const

'Variable
Private WithEvents myComboBox As Access.ComboBox
Attribute myComboBox.VB_VarHelpID = -1

'******************************************************************************************
'*getter/setter
'******************************************************************************************
Public Property Set ComboBox(ByRef cmb As Access.ComboBox)
    Set myComboBox = cmb
    myComboBox.OnKeyDown = "[Event Procedure]"
End Property





'******************************************************************************************
'*Function ÅFdisable keyboard input
'*Arg(1)   ÅFkey code
'*Arg(2)   ÅFshft key pressed or not
'******************************************************************************************
Private Sub myComboBox_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'Const
    Const FUNC_NAME As String = "myComboBox_KeyDown"
    
    'Variable
    
    On Error GoTo ErrorHandler
    
    'prohibit entering (except fror Enter/Tab/Esc)
    If KeyCode = vbKeyReturn Then GoTo ExitHandler
    If KeyCode = vbKeyTab Then GoTo ExitHandler
    If KeyCode = vbKeyEscape Then GoTo ExitHandler
    
    KeyCode = 0
    
    If Not M_EventLog.writeEventLogs(FUNC_NAME) Then GoTo ExitHandler
    
ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Access-Control-WithEvents"
        
    GoTo ExitHandler
        
End Sub

