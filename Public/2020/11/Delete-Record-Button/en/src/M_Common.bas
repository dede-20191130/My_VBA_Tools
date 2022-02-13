Attribute VB_Name = "M_Common"
'@Folder("Database")
Option Compare Database
Option Explicit

'**************************************
'*Common Function Module
'**************************************



'******************************************************************************************
'*Function      :create 10 length random string as a dummy data
'*Arg           :
'*Return        :True > normal termination; False > abnormal termination
'******************************************************************************************
Public Function createDummyString() As String
    
    'Const
    Const FUNC_NAME As String = "createDummyString"
    
    'Variable
    Dim rtnVal As String
    Dim i As Long
    
    On Error GoTo ErrorHandler

    createDummyString = ""

    Call Randomize
    
    rtnVal = String(10, vbNullChar)
    
    For i = 1 To 10
        Mid(rtnVal, InStr(rtnVal, vbNullChar), 1) = Chr(65 + Int(Rnd * (122 - 65 + 1)))
    Next i
    
    createDummyString = rtnVal
    
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

