VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAnalyzeTeam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements clsAbsTeam

'**************************
'*�`�[���N���X�@��̓`�[��
'**************************

'�萔��

'�ϐ���
Private myArrayMenberName(1 To 6) As String



'******************************************************************************************
'*getter/setter��
'******************************************************************************************
Private Property Let clsAbsTeam_arrayMenberName(ByVal idx As Long, ByVal name As String)
    myArrayMenberName(idx) = name
End Property

Private Property Get clsAbsTeam_arrayMenberName(ByVal idx As Long) As String
    clsAbsTeam_arrayMenberName = myArrayMenberName(idx)
End Property



'******************************************************************************************
'*�֐���    �FclsAbsTeam_getMemberName
'*�@�\      �F�`�[�������o�[�̖��O�擾
'*����      �F�Ώێ҂̃C���f�b�N�X�ԍ�
'*�߂�l    �F�`�[�������o�[�̖��O
'******************************************************************************************
Private Function clsAbsTeam_getMemberName(ByVal idx As Long) As String
    
    '�萔
    
    '�ϐ�
    
    clsAbsTeam_getMemberName = idx & "�Ԗڂ̃`�[�������o�[��" & myArrayMenberName(idx) & "�ł��B"
    
    
ExitHandler:

    Exit Function
        
End Function
