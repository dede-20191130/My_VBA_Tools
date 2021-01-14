Attribute VB_Name = "M_Calling"
Option Explicit

'**************************
'*�Ăяo�������W���[��
'**************************


'******************************************************************************************
'*�֐���    �Foutput3rdMemberName
'*�@�\      �F����e�X�g
'*����      �F
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
Public Sub testFunc()
    
    '�萔
    Const FUNC_NAME As String = "testFunc"
    
    '�ϐ�
    Dim team As clsAbsTeam
    Dim coll As New Collection
    
    On Error GoTo ErrorHandler

    '��̓`�[���̖��O��ݒ肷��
    Set team = New clsAnalyzeTeam
    team.arrayMenberName(1) = "����"
    team.arrayMenberName(3) = "Mike"
    team.arrayMenberName(5) = "�뒱"
    
    '��������`�[����ǉ�
    coll.Add team
    coll.Add New clsNewTeam
    
    '���O���o��
    If Not outputSelectedMemberName(coll, 3) Then GoTo ExitHandler

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Interface-Array-Member"
        
    GoTo ExitHandler
        
End Sub



'******************************************************************************************
'*�֐���    �FoutputSelectedMemberName
'*�@�\      �F�����̃R���N�V�����̊e�`�[���́uidx�v�Ԗڂ̃����o�[���O���o�͂���
'*����      �F
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
Private Function outputSelectedMemberName(ByVal collTeam As Collection, ByVal idx As Long) As Boolean
    
    '�萔
    Const FUNC_NAME As String = "outputSelectedMemberName"
    
    '�ϐ�
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

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Interface-Array-Member"
        
    GoTo ExitHandler
        
End Function

