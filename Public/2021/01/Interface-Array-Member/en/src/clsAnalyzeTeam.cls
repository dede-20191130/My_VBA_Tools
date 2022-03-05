VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAnalysisTeam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements clsAbsTeam

'**************************
'*Team Class:  Analysis Teram
'**************************

'Consts

'Vars
Private myArrayMenberName(1 To 6) As String'Max member number of 6



'******************************************************************************************
'*getter/setter
'******************************************************************************************
Private Property Let clsAbsTeam_arrayMenberName(ByVal idx As Long, ByVal name As String)
    myArrayMenberName(idx) = name
End Property

Private Property Get clsAbsTeam_arrayMenberName(ByVal idx As Long) As String
    clsAbsTeam_arrayMenberName = myArrayMenberName(idx)
End Property



'******************************************************************************************
'*Function : get the menber name of index
'*Arg      : index number of target member
'*Return   : the name
'******************************************************************************************
Private Function clsAbsTeam_getMemberName(ByVal idx As Long) As String
    
    'Consts
    
    'Vars
    
    clsAbsTeam_getMemberName = "The " & idx & "th team member is " & myArrayMenberName(idx) 
    
    
ExitHandler: 

    Exit Function
        
End Function

