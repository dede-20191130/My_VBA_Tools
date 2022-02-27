VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_Main 
   Caption         =   "F_Main"
   ClientHeight    =   5220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9330.001
   OleObjectBlob   =   "F_Main.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Explicit


'******************************************************************************************
'*Function :
'*Arg(1)   :
'******************************************************************************************
Private Sub btn_execute_Click()
    
    'Consts
    Const FUNC_NAME As String = "btn_execute_Click"
    
    'Vars
    Dim suffix As String
    Dim objPolymo As clsPolymo
    
    On Error GoTo ErrorHandler
    
    'instantiate a class of processings
    Set objPolymo = New clsPolymo
    
    'get selected processing flag string
    suffix = _
           WorksheetFunction.Rept(Me.rdo_showCurrent.Tag, Abs(CLng(CBool(Me.rdo_showCurrent.Value)))) & _
           WorksheetFunction.Rept(Me.rdo_showUser.Tag, Abs(CLng(CBool(Me.rdo_showUser.Value)))) & _
           WorksheetFunction.Rept(Me.rdo_showGreeting.Tag, Abs(CLng(CBool(Me.rdo_showGreeting.Value))))
    If suffix = "" Then MsgBox "Radio button selection is invalid.", vbCritical, Tool_Name: GoTo ExitHandler
    
    'call corresponding processing function
    If Not CallByName(objPolymo, FUNC_NAME & "_" & suffix, VbMethod) Then GoTo ExitHandler
    

ExitHandler:
    
    Set objPolymo = Nothing
    
    Exit Sub
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, Tool_Name
        
    GoTo ExitHandler
        
End Sub

