VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCmbListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Class")
Option Compare Database
Option Explicit

'**************************
'*�R���{�{�b�N�X�C�x���g���X�i�[
'**************************

'�萔

'�ϐ�
Private WithEvents myComboBox As Access.ComboBox
Attribute myComboBox.VB_VarHelpID = -1

'******************************************************************************************
'*getter/setter
'******************************************************************************************
Public Property Set ComboBox(ByRef cmb As Access.ComboBox)
    Set myComboBox = cmb
    myComboBox.OnKeyDown = "[Event Procedure]"
End Property





'******************************************************************************************
'*�֐���    �FmyComboBox_KeyDown
'*�@�\      �F�L�[���͂𖳌��ɂ���
'*����(1)   �F
'******************************************************************************************
Private Sub myComboBox_KeyDown(KeyCode As Integer, Shift As Integer)
    
    '�萔
    Const FUNC_NAME As String = "myComboBox_KeyDown"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler
    
    '�L�[���͂𖳌��ɂ���(Enter/Tab/Esc�ȊO)
    If KeyCode = vbKeyReturn Then GoTo ExitHandler
    If KeyCode = vbKeyTab Then GoTo ExitHandler
    If KeyCode = vbKeyEscape Then GoTo ExitHandler
    
    KeyCode = 0
    
    If Not M_EventLog.writeEventLogs(FUNC_NAME) Then GoTo ExitHandler
    
ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Access-Control-WithEvents"
        
    GoTo ExitHandler
        
End Sub
