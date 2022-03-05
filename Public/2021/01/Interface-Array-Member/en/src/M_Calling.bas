Attribute VB_Name = "M_Calling"
Option Explicit

'**************************
'*Calling Module
'**************************


'******************************************************************************************
'*FUnction:  operation testing function
'*Arg     :  
'*Return  :  True > normal termination; False > abnormal termination
'******************************************************************************************
Public Sub testFunc()
    
    'Consts
    Const FUNC_NAME As String = "testFunc"
    
    'Vars
    Dim team As clsAbsTeam
    Dim coll As New Collection
    
    On Error GoTo ErrorHandler

    'set names for analysis team members
    Set team = New clsAnalysisTeam
    team.arrayMenberName(1) = "²“¡"
    team.arrayMenberName(3) = "Mike"
    team.arrayMenberName(5) = "Abdallah"
    
    'add analysis team
    'add new team
    coll.Add team
    coll.Add New clsNewTeam
    
    'output 3rd member's name for each team
    If Not outputSelectedMemberName(coll, 3) Then GoTo ExitHandler

ExitHandler: 

    Exit Sub
    
ErrorHandler: 

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name: " & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Interface-Array-Member"
        
    GoTo ExitHandler
        
End Sub



'******************************************************************************************
'*FUnction: Outputs the name of the member whose number is given by index
'*Arg     : collection of the team. All of them implements clsAbsTeam.
'*Arg     : the index number
'*Return  : True > normal termination; False > abnormal termination
'******************************************************************************************
Private Function outputSelectedMemberName(ByVal collTeam As Collection, ByVal idx As Long) As Boolean
    
    'Consts
    Const FUNC_NAME As String = "outputSelectedMemberName"
    
    'Vars
    Dim cntTeam As clsAbsTeam
    
    On Error GoTo ErrorHandler

    outputSelectedMemberName = False
    
    For Each cntTeam In collTeam
        Debug.Print cntTeam.getMemberName(idx)
    Next cntTeam

TruePoint: 

    outputSelectedMemberName = True

ExitHandler: 

    Exit Function
    
ErrorHandler: 

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name: " & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Interface-Array-Member"
        
    GoTo ExitHandler
        
End Function

