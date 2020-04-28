Attribute VB_Name = "M_Create_Estimate_Docs"
'@Folder("VBAProject")
Option Explicit

'******************************************************************************************
'*�֐���    �FGet_Current_Max_Data_Num
'*�@�\      �F���σf�[�^��̍ő�̃f�[�^�ԍ��擾�i�󔒃Z���𖳎��j
'*����(1)   �F
'*�߂�l    �F�ő�̃f�[�^�ԍ�
'******************************************************************************************
Public Function Get_Current_Max_Estimate_Data_Num() As Long
    
    '�萔
    Const FUNC_NAME As String = "Get_Current_Max_Data_Num"
    
    '�ϐ�
    Dim data_num_column_num As Long
    Dim data_num_max_row_cell As Range
    
    On Error GoTo ErrorHandler
    
    '�߂�l�����l
    Get_Current_Max_Estimate_Data_Num = 0
    
    '---�ȉ��ɏ������L�q---
    
    data_num_column_num = ws_estimate_data.Range(STR_NAME_RANGE_ESTIMATE_DATA_HEDD)(1).Column
    
    '���σf�[�^�ԍ���̍ő�s�ԍ��̃f�[�^�i�[�Z�����擾
    Set data_num_max_row_cell = Get_Max_Row_Data_Cell(ws_estimate_data, data_num_column_num)
    
    '���l�ł��邱�Ƃ̒���
    If Not IsNumeric(data_num_max_row_cell.Value) Then
        Err.Raise 1000, "rtn_num", "�y�x���z" & vbTab & "�f�[�^�ԍ��Ƃ��Đ��l���擾�ł��܂���B"
    End If
    
    '�߂�l�ݒ�
    Get_Current_Max_Estimate_Data_Num = data_num_max_row_cell.Value
    
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
'*�֐���    �F�w��̌��Ϗ��p�̏��i�Z�b�g��dict���擾
'*�@�\      �F
'*����(1)   �F�Ώۃf�[�^�ԍ�
'*�߂�l    �Fdict
'******************************************************************************************
Public Function Get_Estimate_Product_Set_Data_Dict_For_Each_Docs() As Object
    
    '�萔
    Const FUNC_NAME As String = "Get_Estimate_Product_Set_Data_Dict_For_Each_Docs"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler
    
    '�߂�l�����l
    Set Get_Estimate_Product_Set_Data_Dict_For_Each_Docs = Nothing
    
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

'******************************************************************************************
'*�֐���    �F�o���f�[�V�����֐��@���C��
'*�@�\      �F�ݒ�l�����������ǂ����`�F�b�N
'*����(1)   �F
'*�߂�l    �F�G���[���e������
'******************************************************************************************
Public Function Is_Valid_Main() As String
    
    '�萔
    Const FUNC_NAME As String = "Is_Valid_Main"
    
    '�ϐ�
    Dim temp_str As String
    Dim rtn_value As String
    
    On Error GoTo ErrorHandler
    
    '�߂�l�����l
    Is_Valid_Main = ""
    
    '---�ȉ��ɏ������L�q---
    '�ݒ藓�̃o���f�[�V����
    temp_str = Is_Valid_Setting_Field
    If temp_str <> "" Then
        rtn_value = rtn_value & vbLf _
                  & vbLf _
                  & temp_str
    End If
    
    '�f�[�^�ԍ��̃o���f�[�V����
    temp_str = Is_Valid_Selected_Data_Num
    If temp_str <> "" Then
        rtn_value = rtn_value & vbLf _
                  & vbLf _
                  & temp_str
    End If
    
    '�߂�l�ݒ�
    Is_Valid_Main = rtn_value
    
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
'*�֐���    �FIs_Valid_Setting_Field
'*�@�\      �F�ݒ藓�̃o���f�[�V����
'*����(1)   �F
'*�߂�l    �F�G���[���e������
'******************************************************************************************
Public Function Is_Valid_Setting_Field() As String
    
    '�萔
    Const FUNC_NAME As String = "Is_Valid_Setting_Field"
    
    '�ϐ�
    Dim rtn_value As String
    
    On Error GoTo ErrorHandler
    
    '�߂�l�����l
    Is_Valid_Setting_Field = ""
    
    '---�ȉ��ɏ������L�q---
    
    '�����
    If Is_Brank_Value(ws_estimate_data.Range(STR_NAME_RANGE_CONSUME_TAX).Value) Then
        rtn_value = Replace(ERR_MSG_INVALID_VALUE, ITEM_KEY_FOR_ERR_MSG, "����ł̒l")
    End If

    '�߂�l�ݒ�
    Is_Valid_Setting_Field = rtn_value
    
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
'*�֐���    �F�f�[�^�ԍ��̃o���f�[�V����
'*�@�\      �Ffrom,to�̐����������Ă��邩�ǂ������`�F�b�N����
'*����(1)   �F
'*�߂�l    �F�G���[���e������
'******************************************************************************************
Public Function Is_Valid_Selected_Data_Num() As String
    
    '�萔
    Const FUNC_NAME As String = "Is_Valid_Selected_Data_Num"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler
    
    '�߂�l�����l
    Is_Valid_Selected_Data_Num = ""
    
    '---�ȉ��ɏ������L�q---
    
    '�󗓃`�F�b�N
    If Is_Brank_Value(F_Create_Estimation_Docs.ComboBox_Target_Num_Start.Value) Or Is_Brank_Value(F_Create_Estimation_Docs.ComboBox_Target_Num_End.Value) Then
        Is_Valid_Selected_Data_Num = Replace(ERR_MSG_INVALID_VALUE, ITEM_KEY_FOR_ERR_MSG, "�쐬�Ώۂ̃f�[�^�ԍ�")
        Exit Function
    End If
    
    '�n�_���I�_�����傫����΃G���[
    If F_Create_Estimation_Docs.ComboBox_Target_Num_Start.Value > F_Create_Estimation_Docs.ComboBox_Target_Num_End.Value Then
        Is_Valid_Selected_Data_Num = Replace(ERR_MSG_INCONSISTENCY_OF_DATA_NUM, ITEM_KEY_FOR_ERR_MSG, "�쐬�Ώۂ̃f�[�^�ԍ�")
        Exit Function
    End If
    
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
'*�֐���    �F���σf�[�^�\�̃o���f�[�V����
'*�@�\      �F�\�̕K�{���ڂ̒l�ُ̈�����m
'*����(1)   �F�Ώۃf�[�^�ԍ�
'*�߂�l(1)    �F�G���[���e������
'*�߂�l(2)    �F���σf�[�^�z��
'******************************************************************************************
Public Function Is_Valid_Estimate_Data_Table(ByVal data_num As Long) As Variant
    
    '�萔
    Const FUNC_NAME As String = "Is_Valid_Estimate_Data_Table"
    
    '�ϐ�
    Dim estimate_data_item_num As Long
    Dim estimate_data_item_rng As Range
    Dim estimate_data_item_rng_1st_cell_row As Long
    Dim estimate_data_item_rng_1st_cell_column As Long
    Dim arr_must_input_data(0 To 6) As Variant
    Dim i As Long
    Dim arr_rtn_value(0 To 1) As Variant
    
    On Error GoTo ErrorHandler
    
    '�߂�l�����l
    arr_rtn_value(0) = ""
    
    '---�ȉ��ɏ������L�q---
100:
    With ws_estimate_data
    
        '�Ώۃf�[�^�s�̍��ڔ͈̓I�u�W�F�N�g
        With .Range(STR_NAME_RANGE_ESTIMATE_DATA_HEDD)
            estimate_data_item_num = .Columns.Count
            estimate_data_item_rng_1st_cell_row = .Item(1).Row
            estimate_data_item_rng_1st_cell_column = .Item(1).Column
        End With
        Set estimate_data_item_rng = .Cells(estimate_data_item_rng_1st_cell_row + data_num, estimate_data_item_rng_1st_cell_column).Resize(, estimate_data_item_num)
    
200:    '# �󗓃`�F�b�N
        '���ϔԍ��`���Ϗ��i�Z�b�g�ԍ��܂ł̒l�����͂���Ă��Ȃ��ƃG���[
        For i = 0 To 6
            arr_must_input_data(i) = estimate_data_item_rng(1, i + 2).Value
        Next i
300:    If Not Is_Not_Brank_Value_For_Array(arr_must_input_data) Then
            arr_rtn_value(0) = Replace(ERR_MSG_ESTIMATE_DATA_MUST_INPUT, ITEM_KEY_FOR_ERR_MSG, "�f�[�^�ԍ�" & data_num)
        End If
        
        '�߂�l�ݒ�
        arr_rtn_value(1) = estimate_data_item_rng.Value
        Is_Valid_Estimate_Data_Table = arr_rtn_value
        
    End With
    
ExitHandler:

    Exit Function
    
ErrorHandler:
    
    If InStr(Err.Description, "�����ꏊ�F") <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        If Err.Number = 9 Then
            Err.Description = ERR_MSG_INVALID_INDEX
        End If
    
        Err.Raise Err.Number, _
                  Err.Source, _
                  "�G���[�ڍׁF" & Err.Description & vbNewLine & _
                  "�����ꏊ�F" & FUNC_NAME & vbNewLine & _
                  "�s�ԍ��F" & Erl & "�i0�͍s�ԍ��ݒ薳���j"
    End If
    
End Function

'******************************************************************************************
'*�֐���    �F���Ϗ��쐬�֐��@���C��
'*�@�\      �F
'*����(1)   �F
'*�߂�l    �F�G���[���e������
'******************************************************************************************
Public Function Create_Estimate_Docs_Main() As String
    
    '�萔
    Const FUNC_NAME As String = "Create_Estimate_Docs_Main"
    Const SHEET_NAME_DUMMY As String = "dummy"
    
    '�ϐ�
    Dim data_num_start As Long
    Dim data_num_end As Long
    Dim data_num_cnt As Long
    Dim new_wb As Workbook
    Dim err_str As String
    Dim currect_time_doc_path_base As String
    
    On Error GoTo ErrorHandler
    '�߂�l�����l
    Create_Estimate_Docs_Main = ""
    
    '---�ȉ��ɏ������L�q---
    
    '�f�[�^�ԍ��̎n�_�ƏI�_���擾
    data_num_start = F_Create_Estimation_Docs.ComboBox_Target_Num_Start.Value
    data_num_end = F_Create_Estimation_Docs.ComboBox_Target_Num_End.Value
    
    '���Ϗ��i�[�p�V�K�u�b�N�쐬
    Set new_wb = Workbooks.Add
    new_wb.Worksheets(1).Name = SHEET_NAME_DUMMY
        
    '�f�[�^�ԍ����ƂɌ��Ϗ��쐬
    For data_num_cnt = data_num_start To data_num_end
        err_str = Create_Estimate_Docs_Each(new_wb, data_num_cnt)
        If err_str <> "" Then
            Create_Estimate_Docs_Main = Create_Estimate_Docs_Main & _
                                        vbLf & _
                                        err_str
        End If
    Next data_num_cnt
    
    '�_�~�[�V�[�g�폜
    If new_wb.Worksheets.Count > 1 Then new_wb.Worksheets(SHEET_NAME_DUMMY).Delete
    
    '�ۑ�
    With CreateObject(STR_ACTIVEX_OBJ_FILE_SYSTEM_OBJ)
        currect_time_doc_path_base = .BuildPath(ThisWorkbook.Path, "���Ϗ�_" & Format(Now, "yyyymmddhhnnss"))
        new_wb.SaveAs currect_time_doc_path_base & ".xlsx"
    End With
    
    'PDF�o��
    If ws_estimate_data.OLEObjects(OP_BUTTON_PDF_EXPORT_ON).Object.Value Then
        Call Create_PDF_Files_For_Estimation_Docs(new_wb, currect_time_doc_path_base)
    End If
    
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
'*�֐���    �F���Ϗ��쐬_�f�[�^�ԍ�����
'*�@�\      �F�f�[�^�ԍ����ƂɃo���f�[�V�����`�F�b�N���A�������쐬����
'*����(1)   �F���Ϗ��V�K�u�b�N
'*����(2)   �F�Ώۃf�[�^�ԍ�
'*�߂�l    �F�G���[���e������
'******************************************************************************************
Public Function Create_Estimate_Docs_Each(ByRef new_wb As Workbook, _
                                          ByVal data_num As Long) As String
    
    '�萔
    Const FUNC_NAME As String = "Create_Estimate_Docs_Each"
    
    '�ϐ�
    Dim temp_arr() As Variant
    Dim rtn_value As String
    Dim arr_estimate_data() As Variant
    Dim arr_set_data_for_data_num As Variant
    Dim arr_used_product_codes() As Variant
    Dim dict_product_data As Object
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    '�߂�l�����l
    Create_Estimate_Docs_Each = ""
    
    '---�ȉ��ɏ������L�q---
    
    '���σf�[�^�\�̃o���f�[�V����
    temp_arr = Is_Valid_Estimate_Data_Table(data_num)
    If temp_arr(0) <> "" Then
        Create_Estimate_Docs_Each = temp_arr(0)
        GoTo ExitHandler
    End If
    'GET�@���σf�[�^�\�̃f�[�^�z��
    arr_estimate_data = temp_arr(1)
    
    'GET�@���Ϗ��i�Z�b�g�f�[�^�z��
    If obj_set_data Is Nothing Then Set obj_set_data = New Cls_Set_Data
    arr_set_data_for_data_num = obj_set_data.Get_Set_Data(arr_estimate_data(1, 8))
    If IsNull(arr_set_data_for_data_num) Then
        Create_Estimate_Docs_Each = Replace(ERR_MSG_ESTIMATE_PRODUCT_SET_DATA_IS_NULL, ITEM_KEY_FOR_ERR_MSG, "�f�[�^�ԍ�" & data_num)
        GoTo ExitHandler
    End If
    
    '�Z�b�g�f�[�^�z�񂩂炷�ׂĂ̏��i�R�[�h�擾
    ReDim arr_used_product_codes(0 To Get_Array_Item_Num(arr_set_data_for_data_num) - 1)
    For i = 0 To Get_Array_Item_Num(arr_set_data_for_data_num) - 1
        arr_used_product_codes(i) = arr_set_data_for_data_num(i)(1, 1)
    Next i
    
    'GET ���i�f�[�^dict
    If obj_product_data Is Nothing Then Set obj_product_data = New Cls_Product_Data
    Set dict_product_data = obj_product_data.Get_Product_Data(arr_used_product_codes)
    
    '���Ϗ��쐬
    Call Insert_Data_To_New_Doc(new_wb, data_num, _
                                arr_estimate_data, _
                                arr_set_data_for_data_num, _
                                dict_product_data)
    
    
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
'*�֐���    �F�V�K���Ϗ��Ƀf�[�^�}��
'*�@�\      �F�����ŗ^����ꂽ�f�[�^�Q��]�L
'*����(1)   �F���Ϗ��V�K�u�b�N
'*����(2)   �F�f�[�^�ԍ�
'*����(3)   �F���σf�[�^�\�̃f�[�^�z��
'*����(4)   �F���Ϗ��i�Z�b�g�f�[�^�z��
'*����(5)   �F���i�f�[�^dict
'******************************************************************************************
Public Sub Insert_Data_To_New_Doc(ByRef new_wb As Workbook, _
                                  ByVal data_num As Long, _
                                  ByVal arr_estimate_data As Variant, _
                                  ByVal arr_set_data_for_data_num As Variant, _
                                  ByVal dict_product_data As Object)
    
    '�萔
    Const FUNC_NAME As String = "Insert_Data_To_New_Doc"
    Const ESTIMATE_SHEET_NAME_BASE As String = "���Ϗ�_No_"
    Const MAX_WRITABLE_ROW_NUM_FOR_PRODUCT_TABLE As Long = 18
    
    '�ϐ�
    Dim ws_target As Worksheet
    Dim product_table_1st_row_num As Long
    Dim product_table_1st_column_num As Long
    Dim product_table_item_num As Long
    Dim i As Long
    Dim cnt_num As Long
    
    On Error GoTo ErrorHandler
    
    '---�ȉ��ɏ������L�q---
    
    '�����e���v���[�g�V�[�g�̃R�s�[
    ThisWorkbook.Worksheets(SHEET_NAME_TEMPLATE).Copy after:=new_wb.Worksheets(new_wb.Worksheets.Count)
    Set ws_target = new_wb.Worksheets(SHEET_NAME_TEMPLATE)
    
    With ws_target
    
        '���l�[��
        .Name = ESTIMATE_SHEET_NAME_BASE & data_num
        
        'SET ���ϓ�
        .Range(STR_NAME_RANGE_TMPL_ESTIMATE_DATE).Value = Replace(.Range(STR_NAME_RANGE_TMPL_ESTIMATE_DATE).Value, "$estimation_day", Format(Now, "yyyy�Nmm��dd��"))
        
        'SET �����
        .Range(STR_NAME_RANGE_TMPL_TAX).Value = ws_estimate_data.Range(STR_NAME_RANGE_CONSUME_TAX).Value
        
        'SET ���σf�[�^
        .Range(STR_NAME_RANGE_TMPL_NUMBER).Value = Replace(.Range(STR_NAME_RANGE_TMPL_NUMBER).Value, "$estimation_serial_num", arr_estimate_data(1, 2))
        .Range(STR_NAME_RANGE_TMPL_COMPANY).Value = Replace(.Range(STR_NAME_RANGE_TMPL_COMPANY).Value, "$company_name", arr_estimate_data(1, 3))
        .Range(STR_NAME_RANGE_TMPL_PERSON_IN_CHARGE).Value = Replace(.Range(STR_NAME_RANGE_TMPL_PERSON_IN_CHARGE).Value, "$parson_in_charge", arr_estimate_data(1, 4))
        .Range(STR_NAME_RANGE_TMPL_DELIVERY_DATE).Value = arr_estimate_data(1, 5)
        .Range(STR_NAME_RANGE_TMPL_PAYMENT_TERMS).Value = arr_estimate_data(1, 6)
        .Range(STR_NAME_RANGE_TMPL_EXPIRATION_DATE).Value = arr_estimate_data(1, 7)
        .Range(STR_NAME_RANGE_TMPL_OTHER_NOTES).Value = arr_estimate_data(1, 9)
        
        '�e�[�u���̃v���p�e�B
        With .Range(STR_NAME_RANGE_PRODUCT_TABLE_HEDD)
            product_table_item_num = .Columns.Count
            product_table_1st_row_num = .Item(1).Row
            product_table_1st_column_num = .Item(1).Column
        End With
        
        'SET ���Ϗ��i�Z�b�g�f�[�^
        cnt_num = 1
        For i = LBound(arr_set_data_for_data_num) To UBound(arr_set_data_for_data_num)
            .Cells(product_table_1st_row_num + cnt_num, product_table_1st_column_num).Value = arr_set_data_for_data_num(i)(1, 1)
            .Cells(product_table_1st_row_num + cnt_num, product_table_1st_column_num + 6).Value = arr_set_data_for_data_num(i)(1, 2)
            .Cells(product_table_1st_row_num + cnt_num, product_table_1st_column_num + 7).Value = arr_set_data_for_data_num(i)(1, 3)
            .Cells(product_table_1st_row_num + cnt_num, product_table_1st_column_num + 12).Value = arr_set_data_for_data_num(i)(1, 4)
            If Not dict_product_data Is Nothing Then
                If dict_product_data.Exists(arr_set_data_for_data_num(i)(1, 1)) Then
                    .Cells(product_table_1st_row_num + cnt_num, product_table_1st_column_num + 3).Value = dict_product_data.Item(arr_set_data_for_data_num(i)(1, 1))(1, 1)
                    .Cells(product_table_1st_row_num + cnt_num, product_table_1st_column_num + 8).Value = dict_product_data.Item(arr_set_data_for_data_num(i)(1, 1))(1, 2)
                End If
            End If
            
            If cnt_num >= 18 Then Exit For
            cnt_num = cnt_num + 1
        Next i
        
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
'*�֐���    �FPDF�쐬
'*�@�\      �F�e���ϕ����V�[�g�ɂ�1�t�@�C����PDF���쐬����
'*����(1)   �F�u�b�N�I�u�W�F�N�g
'*����(2)   �FPDF�i�[�t�H���_�p�X
'******************************************************************************************
Public Sub Create_PDF_Files_For_Estimation_Docs(ByRef new_wb As Workbook, _
                                                ByVal folder_path As String)
    
    '�萔
    Const FUNC_NAME As String = "Create_PDF_Files_For_Estimation_Docs"
    
    '�ϐ�
    Dim cnt_sheet As Variant
    
    On Error GoTo ErrorHandler
    
    '---�ȉ��ɏ������L�q---
    
    'PDF�i�[�t�H���_�쐬
    If Dir(folder_path, vbDirectory) = "" Then MkDir folder_path
    
    'PDF�쐬
    For Each cnt_sheet In new_wb.Worksheets
        If cnt_sheet.Visible Then
            Call cnt_sheet.ExportAsFixedFormat( _
                 Type:=xlTypePDF, _
                 Filename:=folder_path & "\" & cnt_sheet.Name, _
                 IgnorePrintAreas:=False)
        End If
    Next cnt_sheet
    
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


