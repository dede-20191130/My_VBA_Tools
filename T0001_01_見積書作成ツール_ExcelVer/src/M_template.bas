Attribute VB_Name = "M_template"
'******************************************************************************************
'*�֐���    �FXXX
'*�@�\      �F
'*����(1)   �F
'******************************************************************************************
Public Sub XXX()
    
    '�萔
    Const FUNC_NAME As String = "XXX"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler
    '---�ȉ��ɏ������L�q---

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "�G���[���������܂����̂Ń}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ�" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Sub

'******************************************************************************************
'*�֐���    �FYYY
'*�@�\      �F
'*����(1)   �F
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
Public Function YYY() As Boolean
    
    '�萔
    Const FUNC_NAME As String = "YYY"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler
    '�߂�l�����l
    YYY = False
    
    '---�ȉ��ɏ������L�q---


    '�߂�l�ݒ�
    YYY = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "�G���[���������܂����̂Ń}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ�" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Function

'******************************************************************************************
'*�֐���    �FXXX2
'*�@�\      �F
'*����(1)   �F
'******************************************************************************************
Public Sub XXX2()
    
    '�萔
    Const FUNC_NAME As String = "XXX2"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler
    
    '---�ȉ��ɏ������L�q---
    
ExitHandler:

    Exit Sub
    
ErrorHandler:
    
    If InStr(Err.Description, "�����ꏊ�F") <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, _
                  Err.Source, _
                  "�G���[�ڍׁF" & Err.Description & vbNewLine & _
                  "�����ꏊ�F" & FUNC_NAME & vbNewLine & _
                  "�s�ԍ��F" & Erl & "�i0�͍s�ԍ��ݒ薳���j"
    End If
        
End Sub

'******************************************************************************************
'*�֐���    �FYYY2
'*�@�\      �F
'*����(1)   �F
'*�߂�l    �F������
'******************************************************************************************
Public Function YYY2() As String
    
    '�萔
    Const FUNC_NAME As String = "YYY2"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler
    
    '�߂�l�����l
    YYY2 = ""
    
    '---�ȉ��ɏ������L�q---


    '�߂�l�ݒ�
'    YYY2 = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:
    
    If InStr(Err.Description, "�����ꏊ�F") <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, _
                  Err.Source, _
                  "�G���[�ڍׁF" & Err.Description & vbNewLine & _
                  "�����ꏊ�F" & FUNC_NAME & vbNewLine & _
                  "�s�ԍ��F" & Erl & "�i0�͍s�ԍ��ݒ薳���j"
    End If
    
    
End Function


