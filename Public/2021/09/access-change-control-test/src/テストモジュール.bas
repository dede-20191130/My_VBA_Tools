Attribute VB_Name = "�e�X�g���W���[��"
'@Folder("Database")
Option Compare Database
Option Explicit

Private Const SOURCE_NAME = "�e�X�g���W���[��"



'******************************************************************************************
'*�@�\      �F�e�X�g�@�R���{�{�b�N�X�̍��ڃ��X�g��ύX�֐�
'******************************************************************************************
Public Sub �e�X�g_changeCmbBoxItems()
    
    '�萔
    Const FUNC_NAME As String = "�e�X�g_changeCmbBoxItems"
    
    '�ϐ�
    Dim tForm As Form
    Dim fName As String
    Dim cmb As ComboBox
    
    On Error GoTo ErrorHandler

    '//�t�H�[���̓��I�쐬
    Set tForm = CreateForm()
    fName = tForm.Name
    
    '//�f�U�C���r���[�ŊJ��
    DoCmd.OpenForm fName, acDesign
    
    '//�R���{�{�b�N�X�̓��I�쐬
    Set cmb = CreateControl(fName, _
                            AcControlType.acComboBox)
    Dim mycmb As String
    mycmb = "mycmb"
    cmb.Name = mycmb
    cmb.RowSourceType = "Value List"
    
    '//�f�U�C���r���[�����
    DoCmd.Close acForm, fName, acSaveYes
    
    '//�t�H�[���r���[�ŊJ��
    DoCmd.OpenForm fName, acNormal
    
    '//��L�ō쐬�����R���{�{�b�N�X���ēx�Q��
    Set cmb = Forms(fName).Controls(mycmb)
    
    '//���e�X�g01�F�H�ו��̃��X�g�ݒ�
    '////�֐��Ăяo��
    Call changeCmbBoxItems(1, cmb)
    '////�A�T�[�V����
    Debug.Assert cmb.ListCount = 3
    Debug.Assert cmb.Column(0, 0) = "�s�U"
    Debug.Assert cmb.Column(0, 1) = "����"
    Debug.Assert cmb.Column(0, 2) = "�Ă���"
    Debug.Print cmb.ListCount
        
    '//���e�X�g02�F���ݕ��̃��X�g�ݒ�
    '////�֐��Ăяo��
    Call changeCmbBoxItems(2, cmb)
    '////�A�T�[�V����
    Debug.Assert cmb.ListCount = 3
    Debug.Assert cmb.Column(0, 0) = "�R�[��"
    Debug.Assert cmb.Column(0, 1) = "�Β�"
    Debug.Assert cmb.Column(0, 2) = "��"
    Debug.Print cmb.ListCount
    
    '//�t�H�[���r���[�����
    DoCmd.Close , , acSaveNo
    
    '//���I���������t�H�[�����폜
    DoCmd.DeleteObject acForm, fName
    
ExitHandler:
    
    '//�e�X�g����
    Debug.Print Now & ":Finish " & FUNC_NAME
    
    Exit Sub
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�N���X���F" & SOURCE_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.Description, vbCritical
        
    GoTo ExitHandler
        
End Sub



'******************************************************************************************
'*�@�\      �F�e�X�g�@�e�L�X�g�{�b�N�X�̎g�p�\��Ԃ̕ύX�֐�
'******************************************************************************************
Public Sub �e�X�g_changeTextBoxesEnabled()
    
    '�萔
    Const FUNC_NAME As String = "�e�X�g_changeTextBoxesEnabled"
    
    '�ϐ�
    Dim tForm As Form
    Dim fName As String
    Dim textboxes(0 To 3) As textbox
    Dim i As Long
    
    On Error GoTo ErrorHandler

    '//�t�H�[���̓��I�쐬
    Set tForm = CreateForm()
    fName = tForm.Name
    
    '//�f�U�C���r���[�ŊJ��
    DoCmd.OpenForm fName, acDesign
    
    '//�e�L�X�g�{�b�N�X�z��̓��I�쐬
    For i = 0 To 3
        Set textboxes(i) = CreateControl(fName, _
                            AcControlType.acTextBox)
                            
        textboxes(i).Name = "mytext_" & i
        
        '//�ꕔ�̂�under18�A����ȊO��over18�̃^�O��t�^
        If i < 2 Then
            textboxes(i).Tag = "under18"
        Else
            textboxes(i).Tag = "over18"
        End If
    Next i
    
    '//�f�U�C���r���[�����
    DoCmd.Close acForm, fName, acSaveYes
    
    '//�t�H�[���r���[�ŊJ��
    DoCmd.OpenForm fName, acNormal
    
    '//��L�ō쐬�����e�L�X�g�{�b�N�X�z����ēx�Q��
    For i = 0 To 3
        Set textboxes(i) = Forms(fName).Controls("mytext_" & i)
    Next i
    
    '//���e�X�g01�F18�Ζ�����p�̃e�L�X�g�{�b�N�X�̗L����
    '////�֐��Ăяo��
    Call changeTextBoxesEnabled(1, textboxes)
    '////�A�T�[�V����
    Debug.Assert textboxes(0).Tag = "under18"
    Debug.Assert textboxes(0).Enabled = True
    Debug.Assert textboxes(1).Tag = "under18"
    Debug.Assert textboxes(1).Enabled = True
    Debug.Assert textboxes(2).Tag <> "under18"
    Debug.Assert textboxes(2).Enabled = False
    Debug.Assert textboxes(3).Tag <> "under18"
    Debug.Assert textboxes(3).Enabled = False
    
    '//���e�X�g02�F18�Έȏ��p�̃e�L�X�g�{�b�N�X�̗L����
    '////�֐��Ăяo��
    Call changeTextBoxesEnabled(2, textboxes)
    '////�A�T�[�V����
    Debug.Assert textboxes(0).Tag = "under18"
    Debug.Assert textboxes(0).Enabled = False
    Debug.Assert textboxes(1).Tag = "under18"
    Debug.Assert textboxes(1).Enabled = False
    Debug.Assert textboxes(2).Tag <> "under18"
    Debug.Assert textboxes(2).Enabled = True
    Debug.Assert textboxes(3).Tag <> "under18"
    Debug.Assert textboxes(3).Enabled = True
        
    '//�t�H�[���r���[�����
    DoCmd.Close , , acSaveNo
    
    '//���I���������t�H�[�����폜
    DoCmd.DeleteObject acForm, fName
    
ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�N���X���F" & SOURCE_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.Description, vbCritical
        
    GoTo ExitHandler
        
End Sub

