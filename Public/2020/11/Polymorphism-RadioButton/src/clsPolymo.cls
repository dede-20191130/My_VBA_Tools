VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPolymo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Explicit


'******************************************************************************************
'*�֐���    �Fbtn_execute_Click_Current
'*�@�\      �F���ݎ�����\������
'*����(1)   �F
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
Public Function btn_execute_Click_Current() As Boolean
    
    '�萔
    Const FUNC_NAME As String = "btn_execute_Click_Current"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler

    btn_execute_Click_Current = False
    
    '���ݎ�����\������
    MsgBox "���ݎ����F" & Now, , Tool_Name

    btn_execute_Click_Current = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, Tool_Name
        
    GoTo ExitHandler
        
End Function


'******************************************************************************************
'*�֐���    �Fbtn_execute_Click_User
'*�@�\      �F���[�U����\������
'*����(1)   �F
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
Public Function btn_execute_Click_User() As Boolean
    
    '�萔
    Const FUNC_NAME As String = "btn_execute_Click_User"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler

    btn_execute_Click_User = False
    
    With CreateObject("WScript.Network")
        '���[�U����\������
        MsgBox "���[�U���F" & .UserName, , Tool_Name
    End With

    btn_execute_Click_User = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, Tool_Name
        
    GoTo ExitHandler
        
End Function






'******************************************************************************************
'*�֐���    �Fbtn_execute_Click_Greeting
'*�@�\      �F��������\������
'*����(1)   �F
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
Public Function btn_execute_Click_Greeting() As Boolean
    
    '�萔
    Const FUNC_NAME As String = "btn_execute_Click_Greeting"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler

    btn_execute_Click_Greeting = False
    
    MsgBox "����ɂ��́B", , Tool_Name
    
    btn_execute_Click_Greeting = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, Tool_Name
        
    GoTo ExitHandler
        
End Function



