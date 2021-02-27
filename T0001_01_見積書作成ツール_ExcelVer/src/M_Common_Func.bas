Attribute VB_Name = "M_Common_Func"
'@Folder("VBAProject")
Option Explicit



'******************************************************************************************
'*関数名    ：セル入力時ドロップダウンリストをセット
'*機能      ：ドロップダウンリストの更新
'*引数(1)   ：対象のValicationオブジェクト
'*引数(1)   ：セットするリスト文字列
'******************************************************************************************
Public Sub Set_Validation_Dropdown_List(ByRef obj_validation As Validation, _
                                        ByVal dropdown_list As String)
    
    '定数
    Const FUNC_NAME As String = "Set_Validation_Dropdown_List"
    
    '変数
    
    On Error GoTo ErrorHandler
    
    '---以下に処理を記述---
    
    'リストが空ならばカンマのみに置き換える
    If dropdown_list = "" Then dropdown_list = ","
    
    '入力規則にドロップダウンリスト文字列をセット
    With obj_validation
        .Delete
        .Add Type:=xlValidateList, Formula1:=dropdown_list
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
'*関数名    ：配列の要素数取得
'*機能      ：一次元配列なら要素数、二次以上の配列ならば一次元目の要素数を取得する
'*引数(1)   ：配列
'*戻り値    ：要素数
'******************************************************************************************
Public Function Get_Array_Item_Num(ByVal arr As Variant) As Long
    
    '定数
    Const FUNC_NAME As String = "Get_Array_Item_Num"
    
    '変数
    
    On Error GoTo ErrorHandler
    
    '戻り値初期値
    Get_Array_Item_Num = 0
    
    '---以下に処理を記述---
    
    '引数の配列判別
    If Not IsArray(arr) Then GoTo ExitHandler
    
    'Get 要素数
    Get_Array_Item_Num = UBound(arr) - LBound(arr) + 1
    
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
'*関数名    ：Get_Max_Row_Data_Cell
'*機能      ：対象列の最大行番号のデータ格納セルを取得する
'*引数(1)   ：対象シート
'*引数(2)   ：対象列番号
'*戻り値    ：最大行セルオブジェクト
'******************************************************************************************
Public Function Get_Max_Row_Data_Cell(ByVal tgt_sheet As Worksheet, _
                                      ByVal tgt_column_num As Long) As Range
    
    '定数
    Const FUNC_NAME As String = "Get_Max_Row_Data_Cell"
    
    '変数
    Dim tgt_range As Range
    Dim i As Long
    Dim arr_tgt_range_value As Variant
    
    On Error GoTo ErrorHandler
    
    '---以下に処理を記述---
    
    'End関数で下から探索し、探索範囲を指定
    Set tgt_range = tgt_sheet.Range( _
                    tgt_sheet.Cells(1, tgt_column_num), _
                    tgt_sheet.Cells(Rows.Count, tgt_column_num).End(xlUp) _
                    )
    arr_tgt_range_value = tgt_range.Value
    
    '値が空欄であるセルは無視する
    i = tgt_range.Count
    Do
        If i < 2 Then Exit Do
        If Not Is_Brank_Value(arr_tgt_range_value(i, 1)) Then Exit Do
        i = i - 1
    Loop
    
    '戻り値設定
    Set Get_Max_Row_Data_Cell = tgt_range(i)
    
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
'*関数名    ：Delete_Data_Objects
'*機能      ：古いデータ格納オブジェクトの削除
'*引数(1)   ：
'******************************************************************************************
Public Sub Delete_Data_Objects()
    
    '定数
    Const FUNC_NAME As String = "Delete_Data_Objects"
    
    '変数
    
    On Error GoTo ErrorHandler
    
    '---以下に処理を記述---
    
    If Not obj_set_data Is Nothing Then Set obj_set_data = Nothing
    If Not obj_product_data Is Nothing Then Set obj_product_data = Nothing
    
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
'*関数名    ：空欄判別関数
'*機能      ：指定の文字列を空欄かどうか判別する。空白文字を無視する。
'*引数(1)   ：対象文字列
'*戻り値    ：True > 空欄、False > 空欄ではない
'******************************************************************************************
Public Function Is_Brank_Value(ByVal target_str As String) As Boolean
    
    '定数
    Const FUNC_NAME As String = "Is_Brank_Value"
    
    '変数
    Dim rtn_value As Boolean
    
    On Error GoTo ErrorHandler
    
    '戻り値初期値
    Is_Brank_Value = False
    
    '---以下に処理を記述---
    
    rtn_value = (Len( _
                 Replace( _
                 Replace(target_str, " ", ""), _
                 "　", "") _
                 ) = 0)
    

    '戻り値設定
    Is_Brank_Value = rtn_value
    
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
'*関数名    ：配列のコンポーネントに対する空欄判別
'*機能      ：配列中の各値を空欄かどうか判別する。空白文字を無視する。
'*引数(1)   ：対象の配列
'*戻り値    ：True > 空欄の要素が存在しない、False > 空欄が少なくとも1つ存在する
'******************************************************************************************
Public Function Is_Not_Brank_Value_For_Array(ByVal arr As Variant) As Boolean
    
    '定数
    Const FUNC_NAME As String = "Is_Brank_Value_For_Array"
    
    '変数
    Dim i  As Long
    
    On Error GoTo ErrorHandler
    
    '戻り値初期値
    Is_Not_Brank_Value_For_Array = False
    
    '---以下に処理を記述---
    
    '配列であることの判定
    If Not IsArray(arr) Then
        Err.Raise 2000, Err.Source, "【プログラムエラー】配列でない引数が指定されました。"
    End If
    
    '配列の各要素が空欄であるか
    For i = LBound(arr) To UBound(arr)
        If Is_Brank_Value(arr(i)) Then GoTo ExitHandler
    Next i

    '戻り値設定
    Is_Not_Brank_Value_For_Array = True
    
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
'*関数名    ：高速化
'*機能      ：実行速度を高速化する
'*引数(1)   ：描画を省略するフラグ
'*引数(2)   ：手動計算を変更するフラグ
'*引数(3)   ：警告を省略するフラグ
'*引数(4)   ：イベントの発生をコントロールするフラグ
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function Execute_SpeedUp(Optional ByVal is_change_screenUpdating As Boolean = True, _
                                Optional ByVal is_change_calculation As Boolean = True, _
                                Optional ByVal is_change_displayAlerts As Boolean = True, _
                                Optional ByVal is_change_enableEvents As Boolean = True) As Boolean
    
    '定数
    Const FUNC_NAME As String = "Execute_SpeedUp"
    
    '変数
    
    On Error GoTo ErrorHandler
    '戻り値初期値
    Execute_SpeedUp = False
    
    '---以下に処理を記述---

    'マクロ処理中に、描画など余計なものを省略して高速化
    With Application
        If is_change_screenUpdating Then .ScreenUpdating = False '描画を省略
        If is_change_calculation Then .Calculation = xlCalculationManual '手動計算
        If is_change_displayAlerts Then .DisplayAlerts = False '警告を省略。
        If is_change_enableEvents Then .EnableEvents = False 'イベントの発生停止
    End With

    '戻り値設定
    Execute_SpeedUp = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "エラーが発生しましたのでマクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号" & Err.Number & Chr(13) & Err.Description, vbCritical, TOOL_NAME
    Call Reset_SpeedUp
    GoTo ExitHandler
        
End Function

'******************************************************************************************
'*関数名    ：高速化復旧
'*機能      ：高速化で設定したパラメータをもとに戻す
'*引数(1)   ：描画を省略するフラグ
'*引数(2)   ：手動計算を変更するフラグ
'*引数(3)   ：警告を省略するフラグ
'*引数(4)   ：イベントの発生をコントロールするフラグ
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function Reset_SpeedUp(Optional ByVal is_change_screenUpdating As Boolean = True, _
                              Optional ByVal is_change_calculation As Boolean = True, _
                              Optional ByVal is_change_displayAlerts As Boolean = True, _
                              Optional ByVal is_change_enableEvents As Boolean = True) As Boolean
    
    '定数
    Const FUNC_NAME As String = "Reset_SpeedUp"
    
    '変数
    
    On Error Resume Next
    '戻り値初期値
    Reset_SpeedUp = False
    
    '---以下に処理を記述---

    '描画などの設定をリセット
    With Application
        If is_change_screenUpdating Then .ScreenUpdating = True '描画する
        If is_change_calculation Then .Calculation = xlCalculationAutomatic '自動計算
        If is_change_displayAlerts Then .DisplayAlerts = True '警告を行う
        If is_change_enableEvents Then .EnableEvents = True 'イベントの発生
    End With

    '戻り値設定
    Reset_SpeedUp = True
    
ExitHandler:

    Exit Function
    
End Function

'******************************************************************************************
'*関数名    ：Msgboxのラッパー
'*機能      ：高速化に対応
'*引数(1)   ：Msgboxの該当引数
'*引数(2)   ：Msgboxの該当引数
'*引数(3)   ：Msgboxの該当引数
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function Msgbox_Wrapper(ByVal Prompt As String, _
                               Optional ByVal Buttons As VbMsgBoxStyle = VbMsgBoxStyle.vbOKOnly, _
                               Optional ByVal Title As String) As VbMsgBoxResult
    
    '定数
    Const FUNC_NAME As String = "Msgbox_Wrapper"
    
    '変数
    Dim temp_flg As Boolean
    Dim rtn As VbMsgBoxResult
    
    On Error GoTo ErrorHandler
    '戻り値初期値
    Msgbox_Wrapper = VbMsgBoxResult.vbOK
    
    '---以下に処理を記述---
    
    'ScreenUpdatingがFalseならばTrueにする
    temp_flg = Application.ScreenUpdating
    If Not Application.ScreenUpdating Then Application.ScreenUpdating = True
    
    'メッセージ表示
    rtn = MsgBox(Prompt, Buttons, Title)

    '戻り値設定
    Msgbox_Wrapper = rtn
    
ExitHandler:
    
    '復旧
    Application.ScreenUpdating = temp_flg
    
    Exit Function
    
ErrorHandler:
    
    '復旧
    Application.ScreenUpdating = temp_flg
    
    'エラー発生
    On Error GoTo 0
    Err.Raise Err.Number, Err.Source, Err.Description
        
End Function



