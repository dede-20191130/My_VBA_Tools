Attribute VB_Name = "M_Sheet_Event_Triger"
'@Folder("VBAProject")
Option Explicit

'******************************************************************************************
'*関数名    ：Worksheet_Change管理
'*機能      ：シート、セルごとに処理を分ける
'*引数(1)   ：ターゲット範囲オブジェクト
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function Worksheet_Change_Manager(ByRef Target As Range) As Boolean
    
    '定数
    Const FUNC_NAME As String = "Worksheet_Change_Manager"
    
    '変数
    Dim r As Variant
    Dim max_row_data_row_num As Long
    Dim is_trimed As Boolean
    Dim data_rng As Range
    Dim data_rng_val_list_str As String
    Dim cnt As Variant
    Dim new_key As Variant
    Dim skip_flg_product_code_setting As Boolean
    
    On Error GoTo ErrorHandler
    '戻り値初期値
    Worksheet_Change_Manager = False
    
    '---以下に処理を記述---
    
    '高速化、イベント制御
    If Not Execute_SpeedUp() Then GoTo ExitHandler
    
    'シートの分岐
    Select Case Target.Parent.Name
    
    
        '見積商品セットデータシート
    Case SHEET_NAME_ESTIMATE_PRODUCT_SET_DATA
        With ws_estimate_product_set_data
        
            '古いデータ格納オブジェクトの削除
            Call Delete_Data_Objects
            
            'ループ
            For Each r In Target
                'セルの列の分岐
                Select Case r.Column
                    '見積商品セット番号列
                Case .Range(STR_NAME_RANGE_ESTIMATE_PRODUCT_SET_DATA_HEDD)(2).Column
                    If r.Row > .Range(STR_NAME_RANGE_ESTIMATE_PRODUCT_SET_DATA_HEDD)(2).Row Then
                        '見積商品セット番号セル最大行取得
                        If max_row_data_row_num <= 0 Then max_row_data_row_num = Get_Max_Row_Data_Cell(ws_estimate_product_set_data, r.Column).Row
                        If max_row_data_row_num <= 0 Then Exit For
                        If max_row_data_row_num = .Range(STR_NAME_RANGE_ESTIMATE_PRODUCT_SET_DATA_HEDD)(2).Row Then max_row_data_row_num = max_row_data_row_num + 1
                        'データ範囲
                        If data_rng Is Nothing Then Set data_rng = .Range( _
                           .Range(STR_NAME_RANGE_ESTIMATE_PRODUCT_SET_DATA_HEDD)(2).Offset(1, 0), _
                           .Cells(max_row_data_row_num, .Range(STR_NAME_RANGE_ESTIMATE_PRODUCT_SET_DATA_HEDD)(2).Column) _
                           )
                        '見積データの見積商品セット番号リスト更新
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
        
        '商品データシート
    Case SHEET_NAME_PRODUCT_DATA
        With ws_product_data
        
            '古いデータ格納オブジェクトの削除
            Call Delete_Data_Objects
            
            'ループ
            For Each r In Target
                'セルの列の分岐
                Select Case r.Column
                    '商品コード列
                Case .Range(STR_NAME_RANGE_PRODUCT_DATA_HEDD)(2).Column
                    If Not skip_flg_product_code_setting Then
                        If r.Row > .Range(STR_NAME_RANGE_PRODUCT_DATA_HEDD)(2).Row Then
                            '商品データセル最大行取得
                            If max_row_data_row_num <= 0 Then max_row_data_row_num = Get_Max_Row_Data_Cell(ws_product_data, r.Column).Row
                            If max_row_data_row_num <= 0 Then Exit For
                            If max_row_data_row_num = .Range(STR_NAME_RANGE_PRODUCT_DATA_HEDD)(2).Row Then max_row_data_row_num = max_row_data_row_num + 1
                            'データ範囲
                            If data_rng Is Nothing Then Set data_rng = .Range( _
                               .Range(STR_NAME_RANGE_PRODUCT_DATA_HEDD)(2).Offset(1, 0), _
                               .Cells(max_row_data_row_num, .Range(STR_NAME_RANGE_PRODUCT_DATA_HEDD)(2).Column) _
                               )
                               
                            'データ範囲セルの値をトリムする
                            If Not is_trimed Then
                                For Each cnt In data_rng
                                    If cnt.Value <> Trim(cnt.Value) Then cnt.Value = Trim(cnt.Value)
                                Next cnt
                                is_trimed = True
                            End If
                
                            '対象セルの値の重複チェック
                            If WorksheetFunction.CountIf(data_rng, r.Value) > 1 Then
                                MsgBox ERR_MSG_CHANGE_EVENT_DUPLICATION, vbCritical, TOOL_NAME
                                r.Value = ""
                                skip_flg_product_code_setting = True
                                GoTo continue_ws_product_data_1
                            End If
                
                            '見積商品セットデータの商品コードリスト更新
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

    '戻り値設定
    Worksheet_Change_Manager = True
    
ExitHandler:
    
    '高速化、イベント制御
    Call Reset_SpeedUp
    
    Exit Function
    
ErrorHandler:
    
    If Err.Number = 1004 Then
        MsgBox "警告：入力の間隔が短すぎます。", vbExclamation, TOOL_NAME
    Else
        MsgBox "エラーが発生しましたのでマクロを終了します。" & _
               vbLf & _
               "関数名：" & FUNC_NAME & _
               vbLf & _
               "エラー番号" & Err.Number & vbNewLine & _
               Err.Description, vbCritical, TOOL_NAME
        
    End If
    GoTo ExitHandler
        
End Function


