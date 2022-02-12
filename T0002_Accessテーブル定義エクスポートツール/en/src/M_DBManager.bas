Attribute VB_Name = "M_DBManager"
'@Folder("Database")
Option Compare Database
Option Explicit

'**************************
'*DB Management Module
'**************************

'Const


'Variable


'******************************************************************************************
'*Function    :get Access DAO DB instance
'*Arg(1)      :DB file path
'*Return      :True > normal termination; False > abnormal termination
'******************************************************************************************
Public Function getAccessDB(ByVal dbFilePath As String) As DAO.Database
    
    'Const
    Const FUNC_NAME As String = "getAccessDB"
    
    'Variable
    Dim pwStr As String
    Dim errFlg As Boolean
    
    On Error Resume Next
    Set getAccessDB = Nothing
    
    'Open the database
    Set getAccessDB = DBEngine.Workspaces(0).OpenDatabase(dbFilePath, , True)
    
    'exit if not error
    If Err.Number = 0 Then GoTo ExitHandler
    
    'the case of locked with password
    If Err.Number = 3031 Then
        'reset error
        Err.Clear
        'let user input password
        pwStr = InputBox("Please input the password of the Access Database.", "Password Input")
        'open the database again
        Set getAccessDB = DBEngine.Workspaces(0).OpenDatabase(dbFilePath, False, True, "MS Access;PWD=" & pwStr)
        If Err.Number <> 0 Then errFlg = True
    Else
        errFlg = True
    End If
    
    'if error
    If errFlg Then GoTo ErrorHandler
    
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



