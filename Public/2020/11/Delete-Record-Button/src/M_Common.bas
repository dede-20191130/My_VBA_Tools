Attribute VB_Name = "M_Common"
'@Folder("Database")
Option Compare Database
Option Explicit

'**************************************
'*��ʊ֐����W���[��
'**************************************


'�萔

'�ϐ�



'******************************************************************************************
'*�֐���    �FcreateDummyString
'*�@�\      �F�_�~�[�f�[�^�̂��߂�10�����̃����_��������ia~z,A~Z,���̑��L���j�𐶐�
'*����      �F
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
Public Function createDummyString() As String
    
    '�萔
    Const FUNC_NAME As String = "createDummyString"
    
    '�ϐ�
    Dim rtnVal As String
    Dim i As Long
    
    On Error GoTo ErrorHandler

    createDummyString = ""

    Call Randomize
    
    rtnVal = String(10, vbNullChar)
    
    For i = 1 To 10
        Mid(rtnVal, InStr(rtnVal, vbNullChar), 1) = Chr(65 + Int(Rnd * (122 - 65 + 1)))
    Next i
    
    createDummyString = rtnVal
    
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

