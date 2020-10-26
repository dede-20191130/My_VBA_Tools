Attribute VB_Name = "M_DBManager"
'@Folder("Database")
Option Compare Database
Option Explicit

'**************************
'*DB�Ǘ�Module
'**************************

'�萔


'�ϐ�


'******************************************************************************************
'*�֐���    �FgetAccessDB
'*�@�\      �F
'*����(1)   �F
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
Public Function getAccessDB(ByVal dbFilePath As String) As DAO.Database
    
    '�萔
    Const FUNC_NAME As String = "getAccessDB"
    
    '�ϐ�
    Dim pwStr As String
    Dim errFlg As Boolean
    
    On Error Resume Next
    Set getAccessDB = Nothing
    
    '�f�[�^�x�[�X���J��
    Set getAccessDB = DBEngine.Workspaces(0).OpenDatabase(dbFilePath, , True)
    
    '���튮����
    If Err.Number = 0 Then GoTo ExitHandler
    
    '�p�X���[�h���|�����Ă���ꍇ
    If Err.Number = 3031 Then
        '�G���[���Z�b�g
        Err.Clear
        '�p�X���[�h����͂�����
        pwStr = InputBox("Access�f�[�^�x�[�X�̃p�X���[�h����͂��Ă��������B", "�p�X���[�h����")
        '�ēx�f�[�^�x�[�X���J��
        Set getAccessDB = DBEngine.Workspaces(0).OpenDatabase(dbFilePath, False, True, "MS Access;PWD=" & pwStr)
        If Err.Number <> 0 Then errFlg = True
    Else
        errFlg = True
    End If
    
    '�G���[������
    If errFlg Then GoTo ErrorHandler
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.description, vbCritical, Tool_Name
        
    GoTo ExitHandler
        
End Function



