Attribute VB_Name = "mdlMessage"
'@Folder("Database")
Option Compare Database
Option Explicit

'**************************************
'*���b�Z�[�W���W���[��
'**************************************


'�萔

'�ϐ�



'******************************************************************************************
'*�֐���    �FshowErrMessageUpdateRcd
'*�@�\      �F���R�[�h�o�^�E�X�V���G���[���b�Z�[�W
'*����      �F�G���[�v���p�e�B
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
Public Function showErrMessageUpdateRcd(ByVal errNum As Long, ByVal errDescription As String) As Boolean
    
    '�萔
    Const FUNC_NAME As String = "showErrMessageUpdateRcd"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler

    showErrMessageUpdateRcd = False
    
    If errNum = 3022 Then MsgBox "���̒l�͂��łɓo�^�ς݂ł��B", vbCritical, TOOL_NAME: GoTo TruePoint
    
    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & errNum & vbNewLine & _
           errDescription, vbCritical, TOOL_NAME

TruePoint:

    showErrMessageUpdateRcd = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & err.Number & vbNewLine & _
           err.description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Function

