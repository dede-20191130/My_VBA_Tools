Attribute VB_Name = "mdlManageForm"
'@Folder("Database")
Option Compare Database
Option Explicit

'**************************
'*�t�H�[���Ǘ����W���[��
'**************************

'�萔��
Private Const SOURCE_NAME As String = "mdlManageForm"

'�ϐ���



'******************************************************************************************
'*getter/setter��
'******************************************************************************************




'******************************************************************************************
'*�@�\      �F������t�H�[��or�����[�h�t�H�[�����J��
'*����      �F�Ώۃt�H�[�����O
'*����      �F�f�[�^�󂯓n������
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
Public Function showFormInvisibleOrUnloaded(ByVal pFormName As String, Optional pOpenArgs As String = "") As Boolean
    
    '�萔
    Const FUNC_NAME As String = "YshowFormInvisibleOrUnloadedYY"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler

    showFormInvisibleOrUnloaded = False
    
    '���ݓǂݍ��܂�Ă���t�H�[���Ȃ�΍ĉ���
    If Application.CurrentProject.AllForms(pFormName).IsLoaded Then
        Forms(pFormName).Visible = True
    '�t�H�[�����J��
    Else
        DoCmd.OpenForm pFormName, , , , , , _
            pOpenArgs
    End If

TruePoint:

    showFormInvisibleOrUnloaded = True

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




'******************************************************************************************
'*�@�\      �F���[�h�ς݂̃t�H�[�������
'*����      �F�ϒ��@�Ώۂ̃t�H�[��
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
Public Function closeFormIfLoaded(ParamArray pArrFormName() As Variant) As Boolean
    
    '�萔
    Const FUNC_NAME As String = "closeFormIfLoaded"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler

    closeFormIfLoaded = False
    
    Dim c As Variant
    For Each c In pArrFormName
        If Application.CurrentProject.AllForms(CStr(c)).IsLoaded Then DoCmd.Close acForm, CStr(c), acSaveNo
    Next c
    

TruePoint:

    closeFormIfLoaded = True

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

