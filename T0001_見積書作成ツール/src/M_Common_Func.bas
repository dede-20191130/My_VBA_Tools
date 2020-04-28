Attribute VB_Name = "M_Common_Func"
'@Folder("VBAProject")
Option Explicit



'******************************************************************************************
'*�֐���    �F�Z�����͎��h���b�v�_�E�����X�g���Z�b�g
'*�@�\      �F�h���b�v�_�E�����X�g�̍X�V
'*����(1)   �F�Ώۂ�Valication�I�u�W�F�N�g
'*����(1)   �F�Z�b�g���郊�X�g������
'******************************************************************************************
Public Sub Set_Validation_Dropdown_List(ByRef obj_validation As Validation, _
                                        ByVal dropdown_list As String)
    
    '�萔
    Const FUNC_NAME As String = "Set_Validation_Dropdown_List"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler
    
    '---�ȉ��ɏ������L�q---
    
    '���X�g����Ȃ�΃J���}�݂̂ɒu��������
    If dropdown_list = "" Then dropdown_list = ","
    
    '���͋K���Ƀh���b�v�_�E�����X�g��������Z�b�g
    With obj_validation
        .Delete
        .Add Type:=xlValidateList, Formula1:=dropdown_list
    End With
                        
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
'*�֐���    �F�z��̗v�f���擾
'*�@�\      �F�ꎟ���z��Ȃ�v�f���A�񎟈ȏ�̔z��Ȃ�Έꎟ���ڂ̗v�f�����擾����
'*����(1)   �F�z��
'*�߂�l    �F�v�f��
'******************************************************************************************
Public Function Get_Array_Item_Num(ByVal arr As Variant) As Long
    
    '�萔
    Const FUNC_NAME As String = "Get_Array_Item_Num"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler
    
    '�߂�l�����l
    Get_Array_Item_Num = 0
    
    '---�ȉ��ɏ������L�q---
    
    '�����̔z�񔻕�
    If Not IsArray(arr) Then GoTo ExitHandler
    
    'Get �v�f��
    Get_Array_Item_Num = UBound(arr) - LBound(arr) + 1
    
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

'******************************************************************************************
'*�֐���    �FGet_Max_Row_Data_Cell
'*�@�\      �F�Ώۗ�̍ő�s�ԍ��̃f�[�^�i�[�Z�����擾����
'*����(1)   �F�ΏۃV�[�g
'*����(2)   �F�Ώۗ�ԍ�
'*�߂�l    �F�ő�s�Z���I�u�W�F�N�g
'******************************************************************************************
Public Function Get_Max_Row_Data_Cell(ByVal tgt_sheet As Worksheet, _
                                      ByVal tgt_column_num As Long) As Range
    
    '�萔
    Const FUNC_NAME As String = "Get_Max_Row_Data_Cell"
    
    '�ϐ�
    Dim tgt_range As Range
    Dim i As Long
    Dim arr_tgt_range_value As Variant
    
    On Error GoTo ErrorHandler
    
    '---�ȉ��ɏ������L�q---
    
    'End�֐��ŉ�����T�����A�T���͈͂��w��
    Set tgt_range = tgt_sheet.Range( _
                    tgt_sheet.Cells(1, tgt_column_num), _
                    tgt_sheet.Cells(Rows.Count, tgt_column_num).End(xlUp) _
                    )
    arr_tgt_range_value = tgt_range.Value
    
    '�l���󗓂ł���Z���͖�������
    i = tgt_range.Count
    Do
        If i < 2 Then Exit Do
        If Not Is_Brank_Value(arr_tgt_range_value(i, 1)) Then Exit Do
        i = i - 1
    Loop
    
    '�߂�l�ݒ�
    Set Get_Max_Row_Data_Cell = tgt_range(i)
    
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


'******************************************************************************************
'*�֐���    �FDelete_Data_Objects
'*�@�\      �F�Â��f�[�^�i�[�I�u�W�F�N�g�̍폜
'*����(1)   �F
'******************************************************************************************
Public Sub Delete_Data_Objects()
    
    '�萔
    Const FUNC_NAME As String = "Delete_Data_Objects"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler
    
    '---�ȉ��ɏ������L�q---
    
    If Not obj_set_data Is Nothing Then Set obj_set_data = Nothing
    If Not obj_product_data Is Nothing Then Set obj_product_data = Nothing
    
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
'*�֐���    �F�󗓔��ʊ֐�
'*�@�\      �F�w��̕�������󗓂��ǂ������ʂ���B�󔒕����𖳎�����B
'*����(1)   �F�Ώە�����
'*�߂�l    �FTrue > �󗓁AFalse > �󗓂ł͂Ȃ�
'******************************************************************************************
Public Function Is_Brank_Value(ByVal target_str As String) As Boolean
    
    '�萔
    Const FUNC_NAME As String = "Is_Brank_Value"
    
    '�ϐ�
    Dim rtn_value As Boolean
    
    On Error GoTo ErrorHandler
    
    '�߂�l�����l
    Is_Brank_Value = False
    
    '---�ȉ��ɏ������L�q---
    
    rtn_value = (Len( _
                 Replace( _
                 Replace(target_str, " ", ""), _
                 "�@", "") _
                 ) = 0)
    

    '�߂�l�ݒ�
    Is_Brank_Value = rtn_value
    
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

'******************************************************************************************
'*�֐���    �F�z��̃R���|�[�l���g�ɑ΂���󗓔���
'*�@�\      �F�z�񒆂̊e�l���󗓂��ǂ������ʂ���B�󔒕����𖳎�����B
'*����(1)   �F�Ώۂ̔z��
'*�߂�l    �FTrue > �󗓂̗v�f�����݂��Ȃ��AFalse > �󗓂����Ȃ��Ƃ�1���݂���
'******************************************************************************************
Public Function Is_Not_Brank_Value_For_Array(ByVal arr As Variant) As Boolean
    
    '�萔
    Const FUNC_NAME As String = "Is_Brank_Value_For_Array"
    
    '�ϐ�
    Dim i  As Long
    
    On Error GoTo ErrorHandler
    
    '�߂�l�����l
    Is_Not_Brank_Value_For_Array = False
    
    '---�ȉ��ɏ������L�q---
    
    '�z��ł��邱�Ƃ̔���
    If Not IsArray(arr) Then
        Err.Raise 2000, Err.Source, "�y�v���O�����G���[�z�z��łȂ��������w�肳��܂����B"
    End If
    
    '�z��̊e�v�f���󗓂ł��邩
    For i = LBound(arr) To UBound(arr)
        If Is_Brank_Value(arr(i)) Then GoTo ExitHandler
    Next i

    '�߂�l�ݒ�
    Is_Not_Brank_Value_For_Array = True
    
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

'******************************************************************************************
'*�֐���    �F������
'*�@�\      �F���s���x������������
'*����(1)   �F�`����ȗ�����t���O
'*����(2)   �F�蓮�v�Z��ύX����t���O
'*����(3)   �F�x�����ȗ�����t���O
'*����(4)   �F�C�x���g�̔������R���g���[������t���O
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
Public Function Execute_SpeedUp(Optional ByVal is_change_screenUpdating As Boolean = True, _
                                Optional ByVal is_change_calculation As Boolean = True, _
                                Optional ByVal is_change_displayAlerts As Boolean = True, _
                                Optional ByVal is_change_enableEvents As Boolean = True) As Boolean
    
    '�萔
    Const FUNC_NAME As String = "Execute_SpeedUp"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler
    '�߂�l�����l
    Execute_SpeedUp = False
    
    '---�ȉ��ɏ������L�q---

    '�}�N���������ɁA�`��ȂǗ]�v�Ȃ��̂��ȗ����č�����
    With Application
        If is_change_screenUpdating Then .ScreenUpdating = False '�`����ȗ�
        If is_change_calculation Then .Calculation = xlCalculationManual '�蓮�v�Z
        If is_change_displayAlerts Then .DisplayAlerts = False '�x�����ȗ��B
        If is_change_enableEvents Then .EnableEvents = False '�C�x���g�̔�����~
    End With

    '�߂�l�ݒ�
    Execute_SpeedUp = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "�G���[���������܂����̂Ń}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ�" & Err.Number & Chr(13) & Err.Description, vbCritical, TOOL_NAME
    Call Reset_SpeedUp
    GoTo ExitHandler
        
End Function

'******************************************************************************************
'*�֐���    �F����������
'*�@�\      �F�������Őݒ肵���p�����[�^�����Ƃɖ߂�
'*����(1)   �F�`����ȗ�����t���O
'*����(2)   �F�蓮�v�Z��ύX����t���O
'*����(3)   �F�x�����ȗ�����t���O
'*����(4)   �F�C�x���g�̔������R���g���[������t���O
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
Public Function Reset_SpeedUp(Optional ByVal is_change_screenUpdating As Boolean = True, _
                              Optional ByVal is_change_calculation As Boolean = True, _
                              Optional ByVal is_change_displayAlerts As Boolean = True, _
                              Optional ByVal is_change_enableEvents As Boolean = True) As Boolean
    
    '�萔
    Const FUNC_NAME As String = "Reset_SpeedUp"
    
    '�ϐ�
    
    On Error Resume Next
    '�߂�l�����l
    Reset_SpeedUp = False
    
    '---�ȉ��ɏ������L�q---

    '�`��Ȃǂ̐ݒ�����Z�b�g
    With Application
        If is_change_screenUpdating Then .ScreenUpdating = True '�`�悷��
        If is_change_calculation Then .Calculation = xlCalculationAutomatic '�����v�Z
        If is_change_displayAlerts Then .DisplayAlerts = True '�x�����s��
        If is_change_enableEvents Then .EnableEvents = True '�C�x���g�̔���
    End With

    '�߂�l�ݒ�
    Reset_SpeedUp = True
    
ExitHandler:

    Exit Function
    
End Function

'******************************************************************************************
'*�֐���    �FMsgbox�̃��b�p�[
'*�@�\      �F�������ɑΉ�
'*����(1)   �FMsgbox�̊Y������
'*����(2)   �FMsgbox�̊Y������
'*����(3)   �FMsgbox�̊Y������
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
Public Function Msgbox_Wrapper(ByVal Prompt As String, _
                               Optional ByVal Buttons As VbMsgBoxStyle = VbMsgBoxStyle.vbOKOnly, _
                               Optional ByVal Title As String) As VbMsgBoxResult
    
    '�萔
    Const FUNC_NAME As String = "Msgbox_Wrapper"
    
    '�ϐ�
    Dim temp_flg As Boolean
    Dim rtn As VbMsgBoxResult
    
    On Error GoTo ErrorHandler
    '�߂�l�����l
    Msgbox_Wrapper = VbMsgBoxResult.vbOK
    
    '---�ȉ��ɏ������L�q---
    
    'ScreenUpdating��False�Ȃ��True�ɂ���
    temp_flg = Application.ScreenUpdating
    If Not Application.ScreenUpdating Then Application.ScreenUpdating = True
    
    '���b�Z�[�W�\��
    rtn = MsgBox(Prompt, Buttons, Title)

    '�߂�l�ݒ�
    Msgbox_Wrapper = rtn
    
ExitHandler:
    
    '����
    Application.ScreenUpdating = temp_flg
    
    Exit Function
    
ErrorHandler:
    
    '����
    Application.ScreenUpdating = temp_flg
    
    '�G���[����
    On Error GoTo 0
    Err.Raise Err.Number, Err.Source, Err.Description
        
End Function



