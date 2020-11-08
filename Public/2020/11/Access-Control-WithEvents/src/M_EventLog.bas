Attribute VB_Name = "M_EventLog"
'@Folder("Module")
Option Compare Database
Option Explicit


'**************************
'*�C�x���g���OModule
'**************************

'�萔


'�ϐ�
Public targetTxtBox As Access.TextBox


'******************************************************************************************
'*�֐���    �FwriteEventLogs
'*�@�\      �F�e�L�X�g�{�b�N�X�ɃC�x���g���O����������
'*����(1)   �F�L��������
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
Public Function writeEventLogs(ByVal logTxt As String) As Boolean
    
    '�萔
    Const FUNC_NAME As String = "writeEventLogs"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler

    writeEventLogs = False
    
    If targetTxtBox.Value <> "" Then targetTxtBox.Value = targetTxtBox.Value & vbNewLine
    targetTxtBox.Value = targetTxtBox.Value & _
                         Now & _
                         " : " & _
                         logTxt
    
    writeEventLogs = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Access-Control-WithEvents"
        
    GoTo ExitHandler
        
End Function



