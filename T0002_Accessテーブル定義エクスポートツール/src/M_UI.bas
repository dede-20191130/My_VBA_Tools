Attribute VB_Name = "M_UI"
'@Folder("Database")
Option Compare Database
Option Explicit

'**************************
'*���[�U�C���^�t�F�[�XModule
'**************************

'�萔
Public Const msoFileDialogFilePicker As Long = 3


'�ϐ�


'******************************************************************************************
'*�֐���    �FgetFilePathFromDialog
'*�@�\      �F�_�C�A���O�őI�����ꂽ�t�@�C���p�X���擾
'*����(1)   �F�^�C�g��
'*����(2)   �F�t�B���^�Ɏw�肷��key/value�̎���
'*�߂�l    �F�t�@�C���p�X
'******************************************************************************************
Public Function getFilePathFromDialog( _
       Optional ByVal pTitle As String = "�I���_�C�A���O", _
       Optional ByVal dicFilter As Object = Nothing _
       ) As String
    
    '�萔
    Const FUNC_NAME As String = "getFilePathFromDialog"
    
    '�ϐ�
    Dim cntVal As Variant
    Dim filePath As String
    
    On Error GoTo ErrorHandler

    getFilePathFromDialog = ""
    
    '�_�C�A���O�ݒ�
    With Application.FileDialog(msoFileDialogFilePicker)
    
        .Title = pTitle
        
        .Filters.Clear
        If Not dicFilter Is Nothing Then
            For Each cntVal In dicFilter.Keys
                .Filters.Add cntVal, dicFilter.Item(cntVal)
            Next cntVal
            .FilterIndex = 1
        End If
        
        '�����t�@�C���I���̋֎~
        .AllowMultiSelect = False
                
        '�L�����Z�����͏I��
        If .Show <> -1 Then GoTo ExitHandler
        
        '�I�����ꂽ�t�@�C���p�X
        filePath = .SelectedItems(1)
                
    End With

    getFilePathFromDialog = filePath
    
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



