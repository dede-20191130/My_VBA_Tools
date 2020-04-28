VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cls_Product_Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Explicit

Private dict_whole_product_data As Object

'******************************************************************************************
'*関数名    ：コンストラクタ
'*機能      ：
'*引数(1)   ：
'******************************************************************************************
Private Sub Class_Initialize()
    
    '定数
    Const FUNC_NAME As String = "XXX2"
    
    '変数
    Dim product_data_item_num As Long
    Dim product_data_item_rng_1st_cell_row As Long
    Dim product_data_item_rng_1st_cell_column As Long
    Dim max_row_data_row_num As Long
    Dim data_rng As Range
    Dim i As Long
    Dim key_val As Variant
    
    On Error GoTo ErrorHandler
    
    '---以下に処理を記述---
    
    With ws_product_data
        
        'GET ヘッダーrangeの行番号等
        With .Range(STR_NAME_RANGE_PRODUCT_DATA_HEDD)
            product_data_item_num = .Columns.Count
            product_data_item_rng_1st_cell_row = .Item(1).Row
            product_data_item_rng_1st_cell_column = .Item(1).Column
        End With
    
        'GET 商品コードセル最大行
        max_row_data_row_num = Get_Max_Row_Data_Cell(ws_product_data, product_data_item_rng_1st_cell_column + 1).Row
        '記入済みのデータ範囲
        Set data_rng = .Range( _
                       .Range(STR_NAME_RANGE_PRODUCT_DATA_HEDD)(1).Offset(1, 0), _
                       .Cells(max_row_data_row_num, _
                              product_data_item_rng_1st_cell_column + product_data_item_num - 1 _
                              ) _
        )
        
        'SET 商品データdict
        Set dict_whole_product_data = CreateObject(STR_ACTIVEX_OBJ_DICTIONARY)
        For i = 1 To data_rng.Rows.Count
            key_val = data_rng(i, 2).Value
            '登録されていない場合
            If Not dict_whole_product_data.Exists(key_val) Then
                '新規登録
                dict_whole_product_data.Add key_val, .Range(data_rng(i, 3), data_rng(i, product_data_item_num)).Value
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
    Set dict_whole_product_data = Nothing
    
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
'*関数名    ：ゲッター　商品データ
'*機能      ：引数の配列の要素の値に含まれる商品データのdictを取得。すべて存在しなければNothingを返す
'*引数(1)   ：商品コード格納一次元配列
'*戻り値    ：dict
'******************************************************************************************
Public Function Get_Product_Data(ByVal arr As Variant) As Object
    
    '定数
    Const FUNC_NAME As String = "Get_Product_Data"
    
    '変数
    Dim dict_rtn As Object
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    '戻り値初期値
    Set Get_Product_Data = Nothing
    
    '---以下に処理を記述---
    
    '引数の配列判別
    If Not IsArray(arr) Then GoTo ExitHandler
    
    'SET 引数の要素に含まれる商品コードのみの情報保持の新しいdict
    Set dict_rtn = CreateObject(STR_ACTIVEX_OBJ_DICTIONARY)
    For i = LBound(arr) To UBound(arr)
        If dict_whole_product_data.Exists(arr(i)) Then
            dict_rtn.Add arr(i), dict_whole_product_data.Item(arr(i))
        End If
    Next i
    
    '要素数0ならばNothing
    If dict_rtn.Count = 0 Then Set dict_rtn = Nothing

    '戻り値設定
    Set Get_Product_Data = dict_rtn
    
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


