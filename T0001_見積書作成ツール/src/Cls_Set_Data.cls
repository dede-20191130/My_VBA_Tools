VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cls_Set_Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Explicit

Private dict_whole_set_data As Object

'******************************************************************************************
'*関数名    ：コンストラクタ
'*機能      ：セットデータdictを設定する
'*引数(1)   ：
'******************************************************************************************
Private Sub Class_Initialize()
    
    '定数
    Const FUNC_NAME As String = "Class_Initialize"
    
    '変数
    Dim estimate_product_set_data_item_num As Long
    Dim estimate_product_set_data_item_rng_1st_cell_row As Long
    Dim estimate_product_set_data_item_rng_1st_cell_column As Long
    Dim max_row_data_row_num As Long
    Dim data_rng As Range
    Dim i As Long
    Dim key_val As Variant
    
    On Error GoTo ErrorHandler
    '---以下に処理を記述---
    
    With ws_estimate_product_set_data
    
        'Get ヘッダーrangeの行番号等
        With .Range(STR_NAME_RANGE_ESTIMATE_PRODUCT_SET_DATA_HEDD)
            estimate_product_set_data_item_num = .Columns.Count
            estimate_product_set_data_item_rng_1st_cell_row = .Item(1).Row
            estimate_product_set_data_item_rng_1st_cell_column = .Item(1).Column
        End With
    
        'Get 見積商品セット番号セル最大行
        max_row_data_row_num = Get_Max_Row_Data_Cell(ws_estimate_product_set_data, estimate_product_set_data_item_rng_1st_cell_column + 1).Row
        '記入済みのデータ範囲
        Set data_rng = .Range( _
                       .Range(STR_NAME_RANGE_ESTIMATE_PRODUCT_SET_DATA_HEDD)(1).Offset(1, 0), _
                       .Cells(max_row_data_row_num, _
                              estimate_product_set_data_item_rng_1st_cell_column + estimate_product_set_data_item_num - 1 _
                              ) _
        )
        'Set　セットデータdict
        '各セット番号ごとにまとめる
        Set dict_whole_set_data = CreateObject(STR_ACTIVEX_OBJ_DICTIONARY)
        For i = 1 To data_rng.Rows.Count
            key_val = data_rng(i, 2).Value
            'すでに登録されている場合
            If dict_whole_set_data.Exists(key_val) Then
                '商品コードから備考欄までの値の配列を格納
                dict_whole_set_data.Item(key_val).Add dict_whole_set_data.Item(key_val).Count, .Range(data_rng(i, 3), data_rng(i, estimate_product_set_data_item_num)).Value
                '登録されていない場合
            Else
                dict_whole_set_data.Add key_val, CreateObject(STR_ACTIVEX_OBJ_DICTIONARY)
                dict_whole_set_data.Item(key_val).Add dict_whole_set_data.Item(key_val).Count, .Range(data_rng(i, 3), data_rng(i, estimate_product_set_data_item_num)).Value
            End If
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
'*関数名    ：デストラクタ
'*機能      ：
'*引数(1)   ：
'******************************************************************************************
Private Sub Class_Terminate()
    
    '定数
    Const FUNC_NAME As String = "Class_Terminate"
    
    '変数
    
    On Error GoTo ErrorHandler
    
    '---以下に処理を記述---
    
    
ExitHandler:
    
    '解放
    Set dict_whole_set_data = Nothing
    
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
'*関数名    ：ゲッター　セットデータ
'*機能      ：引数のセット番号のデータの配列を取得。存在しなければnullを返す
'*引数(1)   ：セット番号
'*戻り値    ：配列
'******************************************************************************************
Public Function Get_Set_Data(ByVal key_val As Variant) As Variant
    
    '定数
    Const FUNC_NAME As String = "Get_Set_Data"
    
    '変数
    
    On Error GoTo ErrorHandler
    
    '戻り値初期値
    Get_Set_Data = Null
    
    '---以下に処理を記述---
    
    'データ番号がセットデータdict中に存在する場合
    If dict_whole_set_data.Exists(key_val) Then
        Get_Set_Data = dict_whole_set_data.Item(key_val).Items
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


