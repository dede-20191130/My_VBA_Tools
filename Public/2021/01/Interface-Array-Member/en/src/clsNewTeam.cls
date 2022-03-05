VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsNewTeam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements clsAbsTeam

'**************************
'*Team Class:  New Team 
'**************************

'Consts

'Vars


'******************************************************************************************
'*getter/setter
'******************************************************************************************
Private Property Let clsAbsTeam_arrayMenberName(ByVal idx As Long, ByVal name As String)
    'nothing to do
End Property

Private Property Get clsAbsTeam_arrayMenberName(ByVal idx As Long) As String
    clsAbsTeam_arrayMenberName = "There is no member in this new team."
End Property

'******************************************************************************************
'*Function : get the menber name of index, but new team has no member so it returns 'no member'.
'*Arg      : index number of target member
'*Return   : the name
'******************************************************************************************
Private Function clsAbsTeam_getMemberName(ByVal idx As Long) As String
    
    'Consts
    
    'Vars
    
    clsAbsTeam_getMemberName = "There is no member in this new team."
    
    
ExitHandler: 

    Exit Function
        
End Function


