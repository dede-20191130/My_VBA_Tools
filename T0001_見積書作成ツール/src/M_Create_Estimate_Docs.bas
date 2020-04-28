Attribute VB_Name = "M_Create_Estimate_Docs"
'@Folder("VBAProject")
Option Explicit

'******************************************************************************************
'*関数名    ：Get_Current_Max_Data_Num
'*機能      ：見積データ列の最大のデータ番号取得（空白セルを無視）
'*引数(1)   ：
'*戻り値    ：最大のデータ番号
'******************************************************************************************
Public Function Get_Current_Max_Estimate_Data_Num() As Long
    
    '定数
    Const FUNC_NAME As String = "Get_Current_Max_Data_Num"
    
    '変数
    Dim data_num_column_num As Long
    Dim data_num_max_row_cell As Range
    
    On Error GoTo ErrorHandler
    
    '戻り値初期値
    Get_Current_Max_Estimate_Data_Num = 0
    
    '---以下に処理を記述---
    
    data_num_column_num = ws_estimate_data.Range(STR_NAME_RANGE_ESTIMATE_DATA_HEDD)(1).Column
    
    '見積データ番号列の最大行番号のデータ格納セルを取得
    Set data_num_max_row_cell = Get_Max_Row_Data_Cell(ws_estimate_data, data_num_column_num)
    
    '数値であることの調査
    If Not IsNumeric(data_num_max_row_cell.Value) Then
        Err.Raise 1000, "rtn_num", "【警告】" & vbTab & "データ番号として数値を取得できません。"
    End If
    
    '戻り値設定
    Get_Current_Max_Estimate_Data_Num = data_num_max_row_cell.Value
    
ExitHandler:

    Exit Function
    
ErrorHandler:
    
    If InStr(Err.Description, "発生場所：") <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, _
                  Err.Source, _
                  "エラー詳細：" & Err.Description & vbNewLine & _
                  "発生場所：" & FUNC_NAME & vbNewLine & _
                  "行番号：" & Erl & "（0は行番号設定無し）"
    End If
    
End Function

'******************************************************************************************
'*関数名    ：指定の見積書用の商品セットのdictを取得
'*機能      ：
'*引数(1)   ：対象データ番号
'*戻り値    ：dict
'******************************************************************************************
Public Function Get_Estimate_Product_Set_Data_Dict_For_Each_Docs() As Object
    
    '定数
    Const FUNC_NAME As String = "Get_Estimate_Product_Set_Data_Dict_For_Each_Docs"
    
    '変数
    
    On Error GoTo ErrorHandler
    
    '戻り値初期値
    Set Get_Estimate_Product_Set_Data_Dict_For_Each_Docs = Nothing
    
    '---以下に処理を記述---


    '戻り値設定
    '    YYY2 = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:
    
    If InStr(Err.Description, "発生場所：") <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, _
                  Err.Source, _
                  "エラー詳細：" & Err.Description & vbNewLine & _
                  "発生場所：" & FUNC_NAME & vbNewLine & _
                  "行番号：" & Erl & "（0は行番号設定無し）"
    End If
    
    
End Function

'******************************************************************************************
'*関数名    ：バリデーション関数　メイン
'*機能      ：設定値が正しいかどうかチェック
'*引数(1)   ：
'*戻り値    ：エラー内容文字列
'******************************************************************************************
Public Function Is_Valid_Main() As String
    
    '定数
    Const FUNC_NAME As String = "Is_Valid_Main"
    
    '変数
    Dim temp_str As String
    Dim rtn_value As String
    
    On Error GoTo ErrorHandler
    
    '戻り値初期値
    Is_Valid_Main = ""
    
    '---以下に処理を記述---
    '設定欄のバリデーション
    temp_str = Is_Valid_Setting_Field
    If temp_str <> "" Then
        rtn_value = rtn_value & vbLf _
                  & vbLf _
                  & temp_str
    End If
    
    'データ番号のバリデーション
    temp_str = Is_Valid_Selected_Data_Num
    If temp_str <> "" Then
        rtn_value = rtn_value & vbLf _
                  & vbLf _
                  & temp_str
    End If
    
    '戻り値設定
    Is_Valid_Main = rtn_value
    
ExitHandler:

    Exit Function
    
ErrorHandler:
    
    If InStr(Err.Description, "発生場所：") <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, _
                  Err.Source, _
                  "エラー詳細：" & Err.Description & vbNewLine & _
                  "発生場所：" & FUNC_NAME & vbNewLine & _
                  "行番号：" & Erl & "（0は行番号設定無し）"
    End If
        
End Function

'******************************************************************************************
'*関数名    ：Is_Valid_Setting_Field
'*機能      ：設定欄のバリデーション
'*引数(1)   ：
'*戻り値    ：エラー内容文字列
'******************************************************************************************
Public Function Is_Valid_Setting_Field() As String
    
    '定数
    Const FUNC_NAME As String = "Is_Valid_Setting_Field"
    
    '変数
    Dim rtn_value As String
    
    On Error GoTo ErrorHandler
    
    '戻り値初期値
    Is_Valid_Setting_Field = ""
    
    '---以下に処理を記述---
    
    '消費税
    If Is_Brank_Value(ws_estimate_data.Range(STR_NAME_RANGE_CONSUME_TAX).Value) Then
        rtn_value = Replace(ERR_MSG_INVALID_VALUE, ITEM_KEY_FOR_ERR_MSG, "消費税の値")
    End If

    '戻り値設定
    Is_Valid_Setting_Field = rtn_value
    
ExitHandler:

    Exit Function
    
ErrorHandler:
    
    If InStr(Err.Description, "発生場所：") <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, _
                  Err.Source, _
                  "エラー詳細：" & Err.Description & vbNewLine & _
                  "発生場所：" & FUNC_NAME & vbNewLine & _
                  "行番号：" & Erl & "（0は行番号設定無し）"
    End If
    
End Function

'******************************************************************************************
'*関数名    ：データ番号のバリデーション
'*機能      ：from,toの整合性が取れているかどうかをチェックする
'*引数(1)   ：
'*戻り値    ：エラー内容文字列
'******************************************************************************************
Public Function Is_Valid_Selected_Data_Num() As String
    
    '定数
    Const FUNC_NAME As String = "Is_Valid_Selected_Data_Num"
    
    '変数
    
    On Error GoTo ErrorHandler
    
    '戻り値初期値
    Is_Valid_Selected_Data_Num = ""
    
    '---以下に処理を記述---
    
    '空欄チェック
    If Is_Brank_Value(F_Create_Estimation_Docs.ComboBox_Target_Num_Start.Value) Or Is_Brank_Value(F_Create_Estimation_Docs.ComboBox_Target_Num_End.Value) Then
        Is_Valid_Selected_Data_Num = Replace(ERR_MSG_INVALID_VALUE, ITEM_KEY_FOR_ERR_MSG, "作成対象のデータ番号")
        Exit Function
    End If
    
    '始点が終点よりも大きければエラー
    If F_Create_Estimation_Docs.ComboBox_Target_Num_Start.Value > F_Create_Estimation_Docs.ComboBox_Target_Num_End.Value Then
        Is_Valid_Selected_Data_Num = Replace(ERR_MSG_INCONSISTENCY_OF_DATA_NUM, ITEM_KEY_FOR_ERR_MSG, "作成対象のデータ番号")
        Exit Function
    End If
    
ExitHandler:

    Exit Function
    
ErrorHandler:
    
    If InStr(Err.Description, "発生場所：") <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, _
                  Err.Source, _
                  "エラー詳細：" & Err.Description & vbNewLine & _
                  "発生場所：" & FUNC_NAME & vbNewLine & _
                  "行番号：" & Erl & "（0は行番号設定無し）"
    End If
    
End Function

'******************************************************************************************
'*関数名    ：見積データ表のバリデーション
'*機能      ：表の必須項目の値の異常を検知
'*引数(1)   ：対象データ番号
'*戻り値(1)    ：エラー内容文字列
'*戻り値(2)    ：見積データ配列
'******************************************************************************************
Public Function Is_Valid_Estimate_Data_Table(ByVal data_num As Long) As Variant
    
    '定数
    Const FUNC_NAME As String = "Is_Valid_Estimate_Data_Table"
    
    '変数
    Dim estimate_data_item_num As Long
    Dim estimate_data_item_rng As Range
    Dim estimate_data_item_rng_1st_cell_row As Long
    Dim estimate_data_item_rng_1st_cell_column As Long
    Dim arr_must_input_data(0 To 6) As Variant
    Dim i As Long
    Dim arr_rtn_value(0 To 1) As Variant
    
    On Error GoTo ErrorHandler
    
    '戻り値初期値
    arr_rtn_value(0) = ""
    
    '---以下に処理を記述---
100:
    With ws_estimate_data
    
        '対象データ行の項目範囲オブジェクト
        With .Range(STR_NAME_RANGE_ESTIMATE_DATA_HEDD)
            estimate_data_item_num = .Columns.Count
            estimate_data_item_rng_1st_cell_row = .Item(1).Row
            estimate_data_item_rng_1st_cell_column = .Item(1).Column
        End With
        Set estimate_data_item_rng = .Cells(estimate_data_item_rng_1st_cell_row + data_num, estimate_data_item_rng_1st_cell_column).Resize(, estimate_data_item_num)
    
200:    '# 空欄チェック
        '見積番号～見積商品セット番号までの値が入力されていないとエラー
        For i = 0 To 6
            arr_must_input_data(i) = estimate_data_item_rng(1, i + 2).Value
        Next i
300:    If Not Is_Not_Brank_Value_For_Array(arr_must_input_data) Then
            arr_rtn_value(0) = Replace(ERR_MSG_ESTIMATE_DATA_MUST_INPUT, ITEM_KEY_FOR_ERR_MSG, "データ番号" & data_num)
        End If
        
        '戻り値設定
        arr_rtn_value(1) = estimate_data_item_rng.Value
        Is_Valid_Estimate_Data_Table = arr_rtn_value
        
    End With
    
ExitHandler:

    Exit Function
    
ErrorHandler:
    
    If InStr(Err.Description, "発生場所：") <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        If Err.Number = 9 Then
            Err.Description = ERR_MSG_INVALID_INDEX
        End If
    
        Err.Raise Err.Number, _
                  Err.Source, _
                  "エラー詳細：" & Err.Description & vbNewLine & _
                  "発生場所：" & FUNC_NAME & vbNewLine & _
                  "行番号：" & Erl & "（0は行番号設定無し）"
    End If
    
End Function

'******************************************************************************************
'*関数名    ：見積書作成関数　メイン
'*機能      ：
'*引数(1)   ：
'*戻り値    ：エラー内容文字列
'******************************************************************************************
Public Function Create_Estimate_Docs_Main() As String
    
    '定数
    Const FUNC_NAME As String = "Create_Estimate_Docs_Main"
    Const SHEET_NAME_DUMMY As String = "dummy"
    
    '変数
    Dim data_num_start As Long
    Dim data_num_end As Long
    Dim data_num_cnt As Long
    Dim new_wb As Workbook
    Dim err_str As String
    Dim currect_time_doc_path_base As String
    
    On Error GoTo ErrorHandler
    '戻り値初期値
    Create_Estimate_Docs_Main = ""
    
    '---以下に処理を記述---
    
    'データ番号の始点と終点を取得
    data_num_start = F_Create_Estimation_Docs.ComboBox_Target_Num_Start.Value
    data_num_end = F_Create_Estimation_Docs.ComboBox_Target_Num_End.Value
    
    '見積書格納用新規ブック作成
    Set new_wb = Workbooks.Add
    new_wb.Worksheets(1).Name = SHEET_NAME_DUMMY
        
    'データ番号ごとに見積書作成
    For data_num_cnt = data_num_start To data_num_end
        err_str = Create_Estimate_Docs_Each(new_wb, data_num_cnt)
        If err_str <> "" Then
            Create_Estimate_Docs_Main = Create_Estimate_Docs_Main & _
                                        vbLf & _
                                        err_str
        End If
    Next data_num_cnt
    
    'ダミーシート削除
    If new_wb.Worksheets.Count > 1 Then new_wb.Worksheets(SHEET_NAME_DUMMY).Delete
    
    '保存
    With CreateObject(STR_ACTIVEX_OBJ_FILE_SYSTEM_OBJ)
        currect_time_doc_path_base = .BuildPath(ThisWorkbook.Path, "見積書_" & Format(Now, "yyyymmddhhnnss"))
        new_wb.SaveAs currect_time_doc_path_base & ".xlsx"
    End With
    
    'PDF出力
    If ws_estimate_data.OLEObjects(OP_BUTTON_PDF_EXPORT_ON).Object.Value Then
        Call Create_PDF_Files_For_Estimation_Docs(new_wb, currect_time_doc_path_base)
    End If
    
ExitHandler:

    Exit Function
    
ErrorHandler:
    
    If InStr(Err.Description, "発生場所：") <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, _
                  Err.Source, _
                  "エラー詳細：" & Err.Description & vbNewLine & _
                  "発生場所：" & FUNC_NAME & vbNewLine & _
                  "行番号：" & Erl & "（0は行番号設定無し）"
    End If
        
End Function

'******************************************************************************************
'*関数名    ：見積書作成_データ番号ごと
'*機能      ：データ番号ごとにバリデーションチェックし、文書を作成する
'*引数(1)   ：見積書新規ブック
'*引数(2)   ：対象データ番号
'*戻り値    ：エラー内容文字列
'******************************************************************************************
Public Function Create_Estimate_Docs_Each(ByRef new_wb As Workbook, _
                                          ByVal data_num As Long) As String
    
    '定数
    Const FUNC_NAME As String = "Create_Estimate_Docs_Each"
    
    '変数
    Dim temp_arr() As Variant
    Dim rtn_value As String
    Dim arr_estimate_data() As Variant
    Dim arr_set_data_for_data_num As Variant
    Dim arr_used_product_codes() As Variant
    Dim dict_product_data As Object
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    '戻り値初期値
    Create_Estimate_Docs_Each = ""
    
    '---以下に処理を記述---
    
    '見積データ表のバリデーション
    temp_arr = Is_Valid_Estimate_Data_Table(data_num)
    If temp_arr(0) <> "" Then
        Create_Estimate_Docs_Each = temp_arr(0)
        GoTo ExitHandler
    End If
    'GET　見積データ表のデータ配列
    arr_estimate_data = temp_arr(1)
    
    'GET　見積商品セットデータ配列
    If obj_set_data Is Nothing Then Set obj_set_data = New Cls_Set_Data
    arr_set_data_for_data_num = obj_set_data.Get_Set_Data(arr_estimate_data(1, 8))
    If IsNull(arr_set_data_for_data_num) Then
        Create_Estimate_Docs_Each = Replace(ERR_MSG_ESTIMATE_PRODUCT_SET_DATA_IS_NULL, ITEM_KEY_FOR_ERR_MSG, "データ番号" & data_num)
        GoTo ExitHandler
    End If
    
    'セットデータ配列からすべての商品コード取得
    ReDim arr_used_product_codes(0 To Get_Array_Item_Num(arr_set_data_for_data_num) - 1)
    For i = 0 To Get_Array_Item_Num(arr_set_data_for_data_num) - 1
        arr_used_product_codes(i) = arr_set_data_for_data_num(i)(1, 1)
    Next i
    
    'GET 商品データdict
    If obj_product_data Is Nothing Then Set obj_product_data = New Cls_Product_Data
    Set dict_product_data = obj_product_data.Get_Product_Data(arr_used_product_codes)
    
    '見積書作成
    Call Insert_Data_To_New_Doc(new_wb, data_num, _
                                arr_estimate_data, _
                                arr_set_data_for_data_num, _
                                dict_product_data)
    
    
ExitHandler:

    Exit Function
    
ErrorHandler:
    
    If InStr(Err.Description, "発生場所：") <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, _
                  Err.Source, _
                  "エラー詳細：" & Err.Description & vbNewLine & _
                  "発生場所：" & FUNC_NAME & vbNewLine & _
                  "行番号：" & Erl & "（0は行番号設定無し）"
    End If
    
    
End Function

'******************************************************************************************
'*関数名    ：新規見積書にデータ挿入
'*機能      ：引数で与えられたデータ群を転記
'*引数(1)   ：見積書新規ブック
'*引数(2)   ：データ番号
'*引数(3)   ：見積データ表のデータ配列
'*引数(4)   ：見積商品セットデータ配列
'*引数(5)   ：商品データdict
'******************************************************************************************
Public Sub Insert_Data_To_New_Doc(ByRef new_wb As Workbook, _
                                  ByVal data_num As Long, _
                                  ByVal arr_estimate_data As Variant, _
                                  ByVal arr_set_data_for_data_num As Variant, _
                                  ByVal dict_product_data As Object)
    
    '定数
    Const FUNC_NAME As String = "Insert_Data_To_New_Doc"
    Const ESTIMATE_SHEET_NAME_BASE As String = "見積書_No_"
    Const MAX_WRITABLE_ROW_NUM_FOR_PRODUCT_TABLE As Long = 18
    
    '変数
    Dim ws_target As Worksheet
    Dim product_table_1st_row_num As Long
    Dim product_table_1st_column_num As Long
    Dim product_table_item_num As Long
    Dim i As Long
    Dim cnt_num As Long
    
    On Error GoTo ErrorHandler
    
    '---以下に処理を記述---
    
    '文書テンプレートシートのコピー
    ThisWorkbook.Worksheets(SHEET_NAME_TEMPLATE).Copy after:=new_wb.Worksheets(new_wb.Worksheets.Count)
    Set ws_target = new_wb.Worksheets(SHEET_NAME_TEMPLATE)
    
    With ws_target
    
        'リネーム
        .Name = ESTIMATE_SHEET_NAME_BASE & data_num
        
        'SET 見積日
        .Range(STR_NAME_RANGE_TMPL_ESTIMATE_DATE).Value = Replace(.Range(STR_NAME_RANGE_TMPL_ESTIMATE_DATE).Value, "$estimation_day", Format(Now, "yyyy年mm月dd日"))
        
        'SET 消費税
        .Range(STR_NAME_RANGE_TMPL_TAX).Value = ws_estimate_data.Range(STR_NAME_RANGE_CONSUME_TAX).Value
        
        'SET 見積データ
        .Range(STR_NAME_RANGE_TMPL_NUMBER).Value = Replace(.Range(STR_NAME_RANGE_TMPL_NUMBER).Value, "$estimation_serial_num", arr_estimate_data(1, 2))
        .Range(STR_NAME_RANGE_TMPL_COMPANY).Value = Replace(.Range(STR_NAME_RANGE_TMPL_COMPANY).Value, "$company_name", arr_estimate_data(1, 3))
        .Range(STR_NAME_RANGE_TMPL_PERSON_IN_CHARGE).Value = Replace(.Range(STR_NAME_RANGE_TMPL_PERSON_IN_CHARGE).Value, "$parson_in_charge", arr_estimate_data(1, 4))
        .Range(STR_NAME_RANGE_TMPL_DELIVERY_DATE).Value = arr_estimate_data(1, 5)
        .Range(STR_NAME_RANGE_TMPL_PAYMENT_TERMS).Value = arr_estimate_data(1, 6)
        .Range(STR_NAME_RANGE_TMPL_EXPIRATION_DATE).Value = arr_estimate_data(1, 7)
        .Range(STR_NAME_RANGE_TMPL_OTHER_NOTES).Value = arr_estimate_data(1, 9)
        
        'テーブルのプロパティ
        With .Range(STR_NAME_RANGE_PRODUCT_TABLE_HEDD)
            product_table_item_num = .Columns.Count
            product_table_1st_row_num = .Item(1).Row
            product_table_1st_column_num = .Item(1).Column
        End With
        
        'SET 見積商品セットデータ
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
    
    If InStr(Err.Description, "発生場所：") <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, _
                  Err.Source, _
                  "エラー詳細：" & Err.Description & vbNewLine & _
                  "発生場所：" & FUNC_NAME & vbNewLine & _
                  "行番号：" & Erl & "（0は行番号設定無し）"
    End If
        
End Sub

'******************************************************************************************
'*関数名    ：PDF作成
'*機能      ：各見積文書シートにつき1ファイルのPDFを作成する
'*引数(1)   ：ブックオブジェクト
'*引数(2)   ：PDF格納フォルダパス
'******************************************************************************************
Public Sub Create_PDF_Files_For_Estimation_Docs(ByRef new_wb As Workbook, _
                                                ByVal folder_path As String)
    
    '定数
    Const FUNC_NAME As String = "Create_PDF_Files_For_Estimation_Docs"
    
    '変数
    Dim cnt_sheet As Variant
    
    On Error GoTo ErrorHandler
    
    '---以下に処理を記述---
    
    'PDF格納フォルダ作成
    If Dir(folder_path, vbDirectory) = "" Then MkDir folder_path
    
    'PDF作成
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
    
    If InStr(Err.Description, "発生場所：") <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise Err.Number, _
                  Err.Source, _
                  "エラー詳細：" & Err.Description & vbNewLine & _
                  "発生場所：" & FUNC_NAME & vbNewLine & _
                  "行番号：" & Erl & "（0は行番号設定無し）"
    End If
        
End Sub


