Attribute VB_Name = "M_ExceptionHandleTemplate"
Option Explicit

'**************************
'*��O�����T���v��
'**************************




'******************************************************************************************
'*�֐���    �FSub�v���V�[�W���̗�O�����e���v���[�g
'*�@�\      �F
'*����      �F
'******************************************************************************************
Public Sub subTemplate()
    
    '�萔
    Const FUNC_NAME As String = "subTemplate"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler

    '---�������珈�����L�ڂ���---
    

ExitHandler:
    
    '---��������I���������L�ڂ���---
    
    Exit Sub
    
ErrorHandler:
    
    '---���������O�������������L�ڂ���---
    '�@�@��F���b�Z�[�W�{�b�N�X�\���A
    '�@�@�@�@���O�t�@�C���ɃV�X�e���G���[��񏑂����݁A
    '�@�@�@�@�V�X�e���G���[�����̒ʒm���[���̍쐬�E���M�Ȃ�
    
    MsgBox "�V�X�e���G���[���������܂����B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "�}�N��"
        
    GoTo ExitHandler
        
End Sub




'******************************************************************************************
'*�֐���    �FFunction�v���V�[�W���̗�O�����e���v���[�g(1)
'*�@�\      �F
'*����      �F
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
Public Function functionTemplate01() As Boolean
    
    '�萔
    Const FUNC_NAME As String = "functionTemplate01"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler

    functionTemplate01 = False
    
    '---�������珈�����L�ڂ���---

TruePoint:
    
    '---�������琳�펞�݂̂̏I���������L�ڂ���---
    
    functionTemplate01 = True

ExitHandler:
    
    '---��������I���������L�ڂ���---
    
    Exit Function
    
ErrorHandler:

    '---���������O�������������L�ڂ���---
    '�@�@��F���b�Z�[�W�{�b�N�X�\���A
    '�@�@�@�@���O�t�@�C���ɃV�X�e���G���[��񏑂����݁A
    '�@�@�@�@�V�X�e���G���[�����̒ʒm���[���̍쐬�E���M�Ȃ�
    
    MsgBox "�V�X�e���G���[���������܂����B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "�}�N��"
        
    GoTo ExitHandler
        
End Function






'******************************************************************************************
'*�֐���    �FFunction�v���V�[�W���̗�O�����e���v���[�g(2)
'*�@�\      �F
'*����      �F
'*�߂�l    �F�C�ӂ̎w��̊�{�^ > ����I���ANull > �ُ�I��
'******************************************************************************************
Public Function functionTemplate02() As Variant
    
    '�萔
    Const FUNC_NAME As String = "functionTemplate02"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler

    functionTemplate02 = Null
    
    '---�������珈�����L�ڂ���---

ExitHandler:
    
    '---��������I���������L�ڂ���---
    
    Exit Function
    
ErrorHandler:

    '---���������O�������������L�ڂ���---
    '�@�@��F���b�Z�[�W�{�b�N�X�\���A
    '�@�@�@�@���O�t�@�C���ɃV�X�e���G���[��񏑂����݁A
    '�@�@�@�@�V�X�e���G���[�����̒ʒm���[���̍쐬�E���M�Ȃ�
    
    MsgBox "�V�X�e���G���[���������܂����B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "�}�N��"
        
    GoTo ExitHandler
        
End Function

