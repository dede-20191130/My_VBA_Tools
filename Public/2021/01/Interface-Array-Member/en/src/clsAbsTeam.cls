VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAbsTeam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Explicit

'**************************
'*Team Class Interface
'**************************

'Consts

'Vars


'******************************************************************************************
'*getter/setter
'******************************************************************************************
Public Property Let arrayMenberName(ByVal idx As Long, ByVal name As String)

End Property

Public Property Get arrayMenberName(ByVal idx As Long) As String

End Property



'Functions
Public Function getMemberName(ByVal idx As Long) As String

End Function
