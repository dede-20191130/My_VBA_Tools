Attribute VB_Name = "M_Sheet_Event_Triger"
'@Folder("VBAProject")
Option Explicit

'******************************************************************************************
'*�֐���    �FWorksheet_Change�Ǘ�
'*�@�\      �F�V�[�g�A�Z�����Ƃɏ����𕪂���
'*����(1)   �F�^�[�Q�b�g�͈̓I�u�W�F�N�g
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
Public Function Worksheet_Change_Manager(ByRef Target As Range) As Boolean
    
    '�萔
    Const FUNC_NAME As String = "Worksheet_Change_Manager"
    
    '�ϐ�
    Dim r As Variant
    Dim max_row_data_row_num As Long
    Dim is_trimed As Boolean
    Dim data_rng As Range
    Dim data_rng_val_list_str As String
    Dim cnt As Variant
    Dim new_key As Variant
    Dim skip_flg_product_code_setting As Boolean
    
    On Error GoTo ErrorHandler
    '�߂�l�����l
    Worksheet_Change_Manager = False
    
    '---�ȉ��ɏ������L�q---
    
    '�������A�C�x���g����
    If Not Execute_SpeedUp() Then GoTo ExitHandler
    
    '�V�[�g�̕���
    Select Case Target.Parent.Name
    
    
        '���Ϗ��i�Z�b�g�f�[�^�V�[�g
    Case SHEET_NAME_ESTIMATE_PRODUCT_SET_DATA
        With ws_estimate_product_set_data
        
            '�Â��f�[�^�i�[�I�u�W�F�N�g�̍폜
            Call Delete_Data_Objects
            
            '���[�v
            For Each r In Target
                '�Z���̗�̕���
                Select Case r.Column
                    '���Ϗ��i�Z�b�g�ԍ���
                Case .Range(STR_NAME_RANGE_ESTIMATE_PRODUCT_SET_DATA_HEDD)(2).Column
                    If r.Row > .Range(STR_NAME_RANGE_ESTIMATE_PRODUCT_SET_DATA_HEDD)(2).Row Then
                        '���Ϗ��i�Z�b�g�ԍ��Z���ő�s�擾
                        If max_row_data_row_num <= 0 Then max_row_data_row_num = Get_Max_Row_Data_Cell(ws_estimate_product_set_data, r.Column).Row
                        If max_row_data_row_num <= 0 Then Exit For
                        If max_row_data_row_num = .Range(STR_NAME_RANGE_ESTIMATE_PRODUCT_SET_DATA_HEDD)(2).Row Then max_row_data_row_num = max_row_data_row_num + 1
                        '�f�[�^�͈�
                        If data_rng Is Nothing Then Set data_rng = .Range( _
                           .Range(STR_NAME_RANGE_ESTIMATE_PRODUCT_SET_DATA_HEDD)(2).Offset(1, 0), _
                           .Cells(max_row_data_row_num, .Range(STR_NAME_RANGE_ESTIMATE_PRODUCT_SET_DATA_HEDD)(2).Column) _
                           )
                        '���σf�[�^�̌��Ϗ��i�Z�b�g�ԍ����X�g�X�V
                        If data_rng_val_list_str = "" Then
                            With CreateObject(STR_ACTIVEX_OBJ_DICTIONARY)
                                For Each cnt In data_rng
                                    new_key = cnt.Value
                                    If Not .Exists(new_key) Then
                                        .Add new_key, ""
                                    End If
                                Next cnt
                                data_rng_val_list_str = Join(.Keys, ",")
                            End With
                        End If
                        Call Set_Validation_Dropdown_List( _
                             ws_estimate_data.Range(STR_NAME_RANGE_ESTIMATE_DATA_SET_NUM_FIELD).Validation, _
                             data_rng_val_list_str _
                             )
                        
                    End If
                End Select
            Next r
        End With
        
        '���i�f�[�^�V�[�g
    Case SHEET_NAME_PRODUCT_DATA
        With ws_product_data
        
            '�Â��f�[�^�i�[�I�u�W�F�N�g�̍폜
            Call Delete_Data_Objects
            
            '���[�v
            For Each r In Target
                '�Z���̗�̕���
                Select Case r.Column
                    '���i�R�[�h��
                Case .Range(STR_NAME_RANGE_PRODUCT_DATA_HEDD)(2).Column
                    If Not skip_flg_product_code_setting Then
                        If r.Row > .Range(STR_NAME_RANGE_PRODUCT_DATA_HEDD)(2).Row Then
                            '���i�f�[�^�Z���ő�s�擾
                            If max_row_data_row_num <= 0 Then max_row_data_row_num = Get_Max_Row_Data_Cell(ws_product_data, r.Column).Row
                            If max_row_data_row_num <= 0 Then Exit For
                            If max_row_data_row_num = .Range(STR_NAME_RANGE_PRODUCT_DATA_HEDD)(2).Row Then max_row_data_row_num = max_row_data_row_num + 1
                            '�f�[�^�͈�
                            If data_rng Is Nothing Then Set data_rng = .Range( _
                               .Range(STR_NAME_RANGE_PRODUCT_DATA_HEDD)(2).Offset(1, 0), _
                               .Cells(max_row_data_row_num, .Range(STR_NAME_RANGE_PRODUCT_DATA_HEDD)(2).Column) _
                               )
                               
                            '�f�[�^�͈̓Z���̒l���g��������
                            If Not is_trimed Then
                                For Each cnt In data_rng
                                    If cnt.Value <> Trim(cnt.Value) Then cnt.Value = Trim(cnt.Value)
                                Next cnt
                                is_trimed = True
                            End If
                
                            '�ΏۃZ���̒l�̏d���`�F�b�N
                            If WorksheetFunction.CountIf(data_rng, r.Value) > 1 Then
                                MsgBox ERR_MSG_CHANGE_EVENT_DUPLICATION, vbCritical, TOOL_NAME
                                r.Value = ""
                                skip_flg_product_code_setting = True
                                GoTo continue_ws_product_data_1
                            End If
                
                            '���Ϗ��i�Z�b�g�f�[�^�̏��i�R�[�h���X�g�X�V
                            If data_rng_val_list_str = "" Then
                                For Each cnt In data_rng
                                    data_rng_val_list_str = data_rng_val_list_str & _
                                                            cnt.Value & _
                                                            ","
                                Next cnt
                            End If
                            Call Set_Validation_Dropdown_List( _
                                 ws_estimate_product_set_data.Range(STR_NAME_RANGE_ESTIMATE_PRODUCT_SET_DATA_CODE_FIELD).Validation, _
                                 data_rng_val_list_str _
                                 )
                            
                        End If
                    Else
                        r.Value = ""
                    End If
                End Select
continue_ws_product_data_1:
            Next r
        End With
    End Select

    '�߂�l�ݒ�
    Worksheet_Change_Manager = True
    
ExitHandler:
    
    '�������A�C�x���g����
    Call Reset_SpeedUp
    
    Exit Function
    
ErrorHandler:
    
    If Err.Number = 1004 Then
        MsgBox "�x���F���͂̊Ԋu���Z�����܂��B", vbExclamation, TOOL_NAME
    Else
        MsgBox "�G���[���������܂����̂Ń}�N�����I�����܂��B" & _
               vbLf & _
               "�֐����F" & FUNC_NAME & _
               vbLf & _
               "�G���[�ԍ�" & Err.Number & vbNewLine & _
               Err.Description, vbCritical, TOOL_NAME
        
    End If
    GoTo ExitHandler
        
End Function


