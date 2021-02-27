VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_Create_Estimation_Docs 
   Caption         =   "���Ϗ��쐬"
   ClientHeight    =   3435
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7560
   OleObjectBlob   =   "F_Create_Estimation_Docs.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "F_Create_Estimation_Docs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Explicit

'******************************************************************************************
'*�֐���    �FCommandButton_Execute_Create_Click
'*�@�\      �F�Z�Z��
'*����(1)   �F�Z�Z��
'******************************************************************************************
Private Sub CommandButton_Execute_Create_Click()

    '�萔
    Const FUNC_NAME As String = "CommandButton_Execute_Create_Click"
    
    '�ϐ�
    Dim err_str As String
    
    On Error GoTo ErrorHandler
    '---�ȉ��ɏ������L�q---
    
    '������
    If Not Execute_SpeedUp() Then GoTo ExitHandler
    
    '�쐬�O�o���f�[�V�����`�F�b�N
    err_str = Is_Valid_Main
    If err_str <> "" Then
        MsgBox ERR_MSG_CREATED_DOCS_MAIN_HEDD & err_str, vbExclamation, TOOL_NAME
        GoTo ExitHandler
    End If
    
    '���Ϗ��쐬
    err_str = Create_Estimate_Docs_Main
    If err_str <> "" Then
        MsgBox ERR_MSG_CREATED_DOCS_MAIN_EACH_HEDD & vbLf & vbLf & err_str, vbExclamation, TOOL_NAME
    End If
    
    Unload F_Create_Estimation_Docs
    
ExitHandler:
    
    '����
    Call Reset_SpeedUp
    
    Exit Sub
    
ErrorHandler:

    MsgBox "�G���[���������܂����̂Ń}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ�" & Err.Number & Chr(13) & Err.Description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Sub

'******************************************************************************************
'*�֐���    �FUserForm_Initialize
'*�@�\      �F�����������s
'*����(1)   �F
'******************************************************************************************
Private Sub UserForm_Initialize()
    
    '�萔
    Const FUNC_NAME As String = "UserForm_Initialize"
    
    '�ϐ�
    Dim max_data_num As Long
    Dim arr_data_num() As Long
    Dim i As Long
    
    '---�ȉ��ɏ������L�q---
    
    '# �R���{�{�b�N�X�Ƀf�[�^�ԍ����i�[
    '## �f�[�^�ԍ��擾
    max_data_num = Get_Current_Max_Estimate_Data_Num()
    '## �i�[
    ReDim arr_data_num(1 To max_data_num)
    For i = 1 To max_data_num
        arr_data_num(i) = i
    Next i
    ComboBox_Target_Num_Start.List = arr_data_num
    ComboBox_Target_Num_End.List = arr_data_num
    
        
End Sub


