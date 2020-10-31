Attribute VB_Name = "M_UI"
Option Explicit

'**************************
'*���[�U�C���^�t�F�[�XModule
'**************************

'�萔


'�ϐ�


'******************************************************************************************
'*�֐���    �FgetFilePathFromDialog
'*�@�\      �F�_�C�A���O�őI�����ꂽ�t�H���_�p�X���擾
'*����(1)   �F�^�C�g��
'*�߂�l    �F�t�H���_�p�X
'******************************************************************************************
Public Function getFolderPathFromDialog( _
       Optional ByVal pTitle As String = "�I���_�C�A���O" _
       ) As String
    
    '�萔
    Const FUNC_NAME As String = "getFilePathFromDialog"
    
    '�ϐ�
    Dim filePath As String
    
    On Error GoTo ErrorHandler

    getFolderPathFromDialog = ""
    
    '�_�C�A���O�ݒ�
    With Application.FileDialog(msoFileDialogFolderPicker)
    
        .Title = pTitle
        
        '�L�����Z�����͏I��
        If .Show <> -1 Then GoTo ExitHandler
        
        '�I�����ꂽ�t�H���_�p�X
        filePath = .SelectedItems(1)
                
    End With

    getFolderPathFromDialog = filePath
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Function


