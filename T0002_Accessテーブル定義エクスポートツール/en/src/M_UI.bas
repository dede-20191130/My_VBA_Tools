Attribute VB_Name = "M_UI"
'@Folder("Database")
Option Compare Database
Option Explicit

'**************************
'*User Interface Module
'**************************

'Const
Public Const msoFileDialogFilePicker As Long = 3


'Variable


'******************************************************************************************
'*Function    :get the file path picked by dialog
'*Arg(1)      :title
'*Arg(2)      :key-value dictionary object used to filter the files selectable in the dialog
'*Return      :file path

'******************************************************************************************
Public Function getFilePathFromDialog( _
       Optional ByVal pTitle As String = "Pick the target file", _
       Optional ByVal dicFilter As Object = Nothing _
       ) As String
    
    'Const
    Const FUNC_NAME As String = "getFilePathFromDialog"
    
    'Variable
    Dim cntVal As Variant
    Dim filePath As String
    
    On Error GoTo ErrorHandler

    getFilePathFromDialog = ""
    
    'set the dialog
    With Application.FileDialog(msoFileDialogFilePicker)
    
        .Title = pTitle
        
        .Filters.Clear
        If Not dicFilter Is Nothing Then
            For Each cntVal In dicFilter.Keys
                .Filters.Add cntVal, dicFilter.Item(cntVal)
            Next cntVal
            .FilterIndex = 1
        End If
        
        'Disable multiple file selection
        .AllowMultiSelect = False
                
        'exit if canceled
        If .Show <> -1 Then GoTo ExitHandler
        
        filePath = .SelectedItems(1)
                
    End With

    getFilePathFromDialog = filePath
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.description, vbCritical, Tool_Name
        
    GoTo ExitHandler
        
End Function



