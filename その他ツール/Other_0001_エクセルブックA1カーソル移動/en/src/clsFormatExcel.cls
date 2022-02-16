VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormatExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Explicit

'**************************************
'*Class of Moving Cursor to A1 for Excelbook
'**************************************


'Const

'Variable
Private myTgtFileCnt As Long                     'count of target excelbook files


'getter/setter
Public Property Get tgtFileCnt() As Long
    tgtFileCnt = myTgtFileCnt
End Property


#If Not CBool(DEBUG_MODE) Then
Public Property Let tgtFileCnt(ByVal num As Long)
    myTgtFileCnt = num
End Property


#End If


'******************************************************************************************
'******************************************************************************************
Private Sub Class_Initialize()
    
    'Const
    Const FUNC_NAME As String = "Class_Initialize"
    
    'Variable
    
    On Error GoTo ErrorHandler
    
    'pause event execution
    'pause screen drawing
    'omit warnings
    With Application
        .EnableEvents = False
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With

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
Private Sub Class_Terminate()
    
    'Const
    Const FUNC_NAME As String = "Class_Terminate"
    
    'Variable
    
    On Error GoTo ErrorHandler
    
    'restore properties
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With

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
'*Function : execute processing of btn_execute_Click when recursive file exploring is enabled
'*Arg(1)   : folder path
'*Return   : True > normal termination; False > abnormal termination
'******************************************************************************************
Public Function btn_execute_Click_ON(ByVal pFolderPath As String) As Boolean
    
    'Const
    Const FUNC_NAME As String = "btn_execute_Click_ON"
    
    'Variable
    Dim fso As Object
    
    On Error GoTo ErrorHandler

    btn_execute_Click_ON = False
    
    'FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'init file count
    myTgtFileCnt = 0
    
    'explore excel files recursively and process each file
    If Not folderSearchRecurse(pFolderPath, fso) Then GoTo ExitHandler

    btn_execute_Click_ON = True
    
ExitHandler:
    
    Set fso = Nothing
    
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


'******************************************************************************************
'*Function : execute processing of btn_execute_Click when recursive file exploring is not enabled
'*Arg(1)   : folder path
'*Return   : True > normal termination; False > abnormal termination
'******************************************************************************************
Public Function btn_execute_Click_OFF(ByVal pFolderPath As String) As Boolean
    
    'Const
    Const FUNC_NAME As String = "btn_execute_Click_OFF"
    
    'Variable
    Dim fso As Object

    
    On Error GoTo ErrorHandler

    btn_execute_Click_OFF = False
    
    'FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'init file count
    myTgtFileCnt = 0
    
    'explore excel files non-uecursively and process each file
    If Not folderSearch(pFolderPath, fso) Then GoTo ExitHandler

    btn_execute_Click_OFF = True
    
ExitHandler:
    
    Set fso = Nothing
    
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


'******************************************************************************************
'*Function :explore the folder recursively to find Excel files and process each file
'*Arg(1)   :folder path
'*Arg(2)   :FileSystemObject
'*Return   :True > normal termination; False > abnormal termination
'******************************************************************************************
Public Function folderSearchRecurse( _
       ByVal pFolderPath As String, _
       ByVal fso As Object _
       ) As Boolean
    
    'Const
    Const FUNC_NAME As String = "folderSearchRecurse"
    
    'Variable
    Dim folder As Object
    Dim subfolder As Object
    Dim file As Object
  
    On Error GoTo ErrorHandler

    folderSearchRecurse = False
    
    Set folder = fso.GetFolder(pFolderPath)
    
    'list all subfolders in the target folder
    'if there is no subfolder, skip the loop
    For Each subfolder In folder.SubFolders
        'call itself recursively
        If Not folderSearchRecurse(subfolder.Path, fso) Then GoTo ExitHandler
    Next subfolder
  
    'list all files in the target folder
    For Each file In folder.Files
        With file
            'call a cursor moving function
            If Not Me.formatExcel(file.Path, fso) Then GoTo ExitHandler
        End With
    Next file


    folderSearchRecurse = True
    
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




'******************************************************************************************
'*Function :explore the folder to find Excel files and process each file
'*Arg(1)   :folder path
'*Arg(2)   :FileSystemObject
'*Return   :True > normal termination; False > abnormal termination
'******************************************************************************************
Public Function folderSearch( _
       ByVal pFolderPath As String, _
       ByVal fso As Object _
       ) As Boolean
    
    'Const
    Const FUNC_NAME As String = "folderSearch"
    
    'Variable
    Dim folder As Object
    Dim subfolder As Object
    Dim file As Object
  
    On Error GoTo ErrorHandler

    folderSearch = False
    
    Set folder = fso.GetFolder(pFolderPath)
  
    'list all files in the target folder
    For Each file In folder.Files
        With file
            'call a cursor moving function
            If Not Me.formatExcel(file.Path, fso) Then GoTo ExitHandler
        End With
    Next file


    folderSearch = True
    
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




'******************************************************************************************
'*Function :format the Excelbook
'           This move the cursor of each sheet to the A1 Cell and select the leftmost sheet.
'*Arg(1)   :file path
'*Arg(2)   :FileSystemObject
'*Return   :True > normal termination; False > abnormal termination
'******************************************************************************************
Public Function formatExcel( _
       ByVal pFilePath As String, _
       ByVal fso As Object _
       ) As Boolean
    
    'Const
    Const FUNC_NAME As String = "formatExcel"
    
    'Variable
    Dim extension As String
    Dim wb As Workbook
    Dim cntObj As Worksheet
    
    On Error GoTo ErrorHandler

    formatExcel = False
    
    'exit if its extension is not one of those of Excel file
    extension = fso.GetExtensionName(pFilePath)
    If Not ( _
       extension = "xls" Or _
       extension = "xlsx" Or _
       extension = "xlsm" _
       ) Then formatExcel = True: GoTo ExitHandler
    
    'exit if target file is a temporary file
    If Left(fso.GetFileName(pFilePath), 2) = "~$" Then formatExcel = True: GoTo ExitHandler
    
    'open the book
    'exit if the file can't be opened as Excel
    On Error Resume Next
    Set wb = Application.Workbooks.Open(pFilePath)
    If Err.Number <> 0 Then Err.Clear: formatExcel = True: GoTo ExitHandler
    Err.Clear
    On Error GoTo ErrorHandler
    
    With wb
        'move the cursor to A1 Cell
        For Each cntObj In .Worksheets
            If cntObj.Visible = True Then
                cntObj.Select
                cntObj.Range("A1").Select
                'change the position of scrolls
                With ActiveWindow
                    .ScrollRow = ActiveCell.Row
                    .ScrollColumn = ActiveCell.Column
                End With
            End If
        Next
    
        'select the leftmost visible sheet
        For Each cntObj In .Worksheets
            If cntObj.Visible = True Then
                cntObj.Select
                Exit For
            End If
        Next
    End With
    
    'Save
    wb.Save
    
    'increment the count
    myTgtFileCnt = myTgtFileCnt + 1
    
    formatExcel = True
    
ExitHandler:
    
    If Not wb Is Nothing Then wb.Close False: Set wb = Nothing
    
    Set cntObj = Nothing
    
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



