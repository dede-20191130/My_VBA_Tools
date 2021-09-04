Attribute VB_Name = "�R���g���[������n�֐����W���[��"
'@Folder("Database")
Option Compare Database
Option Explicit

Private Const SOURCE_NAME = "�R���g���[������n�֐����W���[��"

'******************************************************************************************
'*�@�\      �F�R���{�{�b�N�X�̍��ڃ��X�g��ύX
'*����      �F
'******************************************************************************************
Public Sub changeCmbBoxItems(ByVal selectedNumber As Long, ByVal cmbBox As ComboBox)
    
    '�萔
    Const FUNC_NAME As String = "changeCmbBoxItems"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler
    
    '//���ڂ̃N���A
    cmbBox.RowSource = ""
    
    Select Case selectedNumber
    '//�H�ו�
    Case 1
        cmbBox.AddItem "�s�U"
        cmbBox.AddItem "����"
        cmbBox.AddItem "�Ă���"
    '//���ݕ�
    Case 2
        cmbBox.AddItem "�R�[��"
        cmbBox.AddItem "�Β�"
        cmbBox.AddItem "��"
    End Select

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





'******************************************************************************************
'*�@�\      �F�e�L�X�g�{�b�N�X�̎g�p�\��Ԃ�ύX
'*����      �F
'******************************************************************************************
Public Sub changeTextBoxesEnabled(ByVal selectedNumber As Long, ByRef textboxes() As textbox)
    
    '�萔
    Const FUNC_NAME As String = "changeTextBoxesEnabled"
    
    '�ϐ�
    Dim canUnder18Enable As Boolean '//18�Ζ����̂��߂̃e�L�X�g�{�b�N�X���L�����ǂ���
    Dim textbox As Variant
    
    On Error GoTo ErrorHandler
    
    '//18�Ζ�����I������True�A����ȊO�̏ꍇ��False
    canUnder18Enable = selectedNumber = 1
    
    '//�^�O��under18��over18���ɂ����
    '//�g�p�\��Ԃ�؂�ւ���
    For Each textbox In textboxes
        If InStr(textbox.Tag, "under18") <> 0 Then
            textbox.Enabled = canUnder18Enable
        Else
            textbox.Enabled = Not canUnder18Enable
        End If
    Next textbox

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


