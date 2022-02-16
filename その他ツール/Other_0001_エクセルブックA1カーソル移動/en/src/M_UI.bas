Attribute VB_Name = "M_UI"
Option Explicit

'**************************
'*User Interface Module
'**************************

'Const


'Variable


'******************************************************************************************
'*Function :get the folder path selected in a dialog
'*Arg(1)   :title string
'*Return   :folder path
'******************************************************************************************
Public Function getFolderPathFromDialog( _
       Optional ByVal pTitle As String = "SELECTION DIALOG" _
       ) As String
    
    'Const
    Const FUNC_NAME As String = "getFilePathFromDialog"
    
    'Variable
    Dim folderPath As String
    
    On Error GoTo ErrorHandler

    getFolderPathFromDialog = ""
    
    'set the dialog
    With Application.FileDialog(msoFileDialogFolderPicker)
    
        .Title = pTitle
        
        'exit if cancelled
        If .Show <> -1 Then GoTo ExitHandler
        
        'folder path
        folderPath = .SelectedItems(1)
                
    End With

    getFolderPathFromDialog = folderPath
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Function


