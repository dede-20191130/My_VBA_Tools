VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_Main 
   Caption         =   "Other_0001 Move Cursor to A1 in Excelbook"
   ClientHeight    =   4920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9780.001
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
'*Function :terminate the tool when closing the form
'******************************************************************************************
Private Sub UserForm_Terminate()

    
    'Const
    Const FUNC_NAME As String = "UserForm_Terminate"
    
    'Variable
    
    On Error GoTo ErrorHandler
    
    #If Not CBool(DEBUG_MODE) Then
        ThisWorkbook.Close False
    #End If
    
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
Private Sub btn_execute_Click()
    
    'Const
    Const FUNC_NAME As String = "btn_execute_Click"
    
    'Variable
    Dim suffix As String
    Dim folderPath As String
    Dim objFormatExcel As clsFormatExcel
    
    On Error GoTo ErrorHandler
    
    Set objFormatExcel = New clsFormatExcel
    
    'get folder path via dialog
    folderPath = getFolderPathFromDialog("SPECIFY TARGT FOLDER")
    If folderPath = "" Then GoTo ExitHandler
        
    'check if recursive file exploration is valid
    suffix = _
           WorksheetFunction.Rept(Me.opt_onRecurse.Tag, Abs(CLng(CBool(Me.opt_onRecurse.Value)))) & _
           WorksheetFunction.Rept(Me.opt_offRecurse.Tag, Abs(CLng(CBool(Me.opt_offRecurse.Value))))
    If suffix = "" Then GoTo ExitHandler
        
    'call a processing function for each condition (suffix)
    If Not CallByName(objFormatExcel, FUNC_NAME & "_" & suffix, VbMethod, folderPath) Then GoTo ExitHandler
        
    'notice the processed file count
    MsgBox objFormatExcel.tgtFileCnt & " Excel files have been cursor-moved.", , TOOL_NAME

ExitHandler:
    
    Set objFormatExcel = Nothing
    
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
Private Sub btn_terminate_Click()
    
    'Const
    Const FUNC_NAME As String = "btn_terminate_Click"
    
    'Variable
    
    On Error GoTo ErrorHandler
    
    'message
    If MsgBox("The tool will be exited.", vbYesNo, TOOL_NAME) <> vbYes Then GoTo ExitHandler
    
    Unload Me
    

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




