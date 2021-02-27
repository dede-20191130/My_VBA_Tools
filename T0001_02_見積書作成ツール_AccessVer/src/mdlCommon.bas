Attribute VB_Name = "mdlCommon"
'@Folder("Database")
Option Compare Database
Option Explicit

'**************************************
'*��ʊ֐����W���[��
'**************************************


'�萔

'�ϐ�




'******************************************************************************************
'*�֐���    �FinitializeTool
'*�@�\      �F
'*����(1)   �F
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
Public Function initializeTool() As Boolean
    
    '�萔
    Const FUNC_NAME As String = "initializeTool"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler

    initializeTool = False
    
    '�O���[�o���ϐ��̐ݒ�
    Set gObjDB = New clsDB
    Set gObjDtTrsfrManager = New clsDtTrsfrManager

    initializeTool = True
    
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
'*�@�\      �FCollection�̃A�C�e����z��ɕϊ�
'*����      �F�Ώ�
'*�߂�l    �F�z��
'******************************************************************************************
Public Function CollectionToArray(myCol As Collection) As Variant
    
    '�萔
    Const FUNC_NAME As String = "CollectionToArray"
    
    '�ϐ�
    Dim result  As Variant
    Dim cnt     As Long
    
    CollectionToArray = False
    
    ReDim result(0 To myCol.Count - 1)

    For cnt = 0 To myCol.Count - 1
        result(cnt) = myCol(cnt + 1)
    Next cnt

    CollectionToArray = result


ExitHandler:

    Exit Function
    
End Function

