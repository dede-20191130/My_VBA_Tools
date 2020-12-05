Attribute VB_Name = "M_Test"
Option Explicit



'**************************
'*例外処理サンプル
'**************************






'******************************************************************************************
'*関数名    ：例外処理Subプロシージャ実例
'*機能      ：
'*引数      ：
'******************************************************************************************
Public Sub subSample()
    
    '定数
    Const FUNC_NAME As String = "subSample"
    
    '変数
    Dim filePathArr As Variant
    Dim filePath As Variant
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    'funcSample01の呼び出し　存在しないシート名の引数で呼び出す
    filePathArr = funcSample01("sheetNotExist")
    '戻り値がNullであるためメッセージ表示
    If IsNull(filePathArr) Then MsgBox "ファイルパス配列の取得に失敗しました（処理は続行します）。"
    
    'funcSample01の呼び出し　存在するシート名の引数で呼び出す
    filePathArr = funcSample01("FilePath")
    '戻り値がNullではないため失敗の表示なし
    If IsNull(filePathArr) Then MsgBox "ファイルパス配列の取得に失敗しました（処理は続行します）。"
    
    'それぞれのエクセルファイルについて、funcSample02を呼び出す
    For Each filePath In filePathArr
        'funcSample02の呼び出し
        'すでにA1セルが書き込まれていた場合はイミディエイトウィンドウに失敗したファイルパスを出力
        If Not funcSample02(ThisWorkbook.Path & filePath) Then
            Debug.Print "書き込み失敗ファイル：" & filePath
        End If
    Next filePath
    
    
    '■■■funcSample01,funcSample02などでキャッチできなかった想定外のエラーは
    '　　　このプロシージャのErrorHandlerでキャッチされます。
    
ExitHandler:
    
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrorHandler:

    MsgBox "システムエラーが発生しました。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "マクロ"
        
    GoTo ExitHandler
        
End Sub




'******************************************************************************************
'*関数名    ：例外処理Functionプロシージャ実例(1)
'*機能      ：ファイルパスの文字列を配列として取得
'*引数      ：このファイルのシートの名前
'*戻り値    ：文字列の配列 > 正常終了、Null > 異常終了
'******************************************************************************************
Public Function funcSample01(ByVal wsName As String) As Variant
    
    '定数
    Const FUNC_NAME As String = "funcSample01"
    
    '変数
    
    On Error GoTo ErrorHandler

    funcSample01 = Null
    
    '指定されたシートのA1セルからA3セルまでの値を配列として取得する
    With ThisWorkbook.Worksheets(wsName)
        funcSample01 = .Range("A1:A3").Value
    End With

ExitHandler:
    
    
    Exit Function
    
ErrorHandler:

    MsgBox "システムエラーが発生しました。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "マクロ"
        
    GoTo ExitHandler
        
End Function



'******************************************************************************************
'*関数名    ：例外処理Functionプロシージャ実例(2)
'*機能      ：指定されたパスのエクセルファイルを開く
'               一枚目のシートのA1セルに時刻を書き込む
'               二枚目のシートが存在すれば、二枚目のA1セルに「完了」と書き込む
'*引数      ：エクセルファイルのパス
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function funcSample02(ByVal filePath As String) As Boolean
    
    '定数
    Const FUNC_NAME As String = "funcSample02"
    
    '変数
    Dim wb As Workbook
    
    On Error GoTo ErrorHandler

    funcSample02 = False
    
    Set wb = Workbooks.Open(filePath)
    
    
    With wb
        '一枚目のシートのA1セルに時刻を書き込む
        'すでにA1セルに文字が書き込まれていた場合はエラーとなる（異常終了）
        If Trim(.Worksheets(1).Range("A1").Value) <> "" Then Err.Raise 1000, , "A1セルにすでに値が存在します。"
        .Worksheets(1).Range("A1").Value = Now
        
        '二枚目のシートが存在しなければ終了（正常終了）
        If .Worksheets.Count < 2 Then GoTo TruePoint
        
        '二枚目のA1セルに「完了」と書き込む
        .Worksheets(2).Range("A1").Value = "完了"
        
    End With
    

TruePoint:
    
    'シートの保存
    wb.Save
    
    funcSample02 = True

ExitHandler:
    
    '正常終了時でもエラーが起きた場合でも、必ずブックを閉じる
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    
    Exit Function
    
ErrorHandler:

    MsgBox "システムエラーが発生しました。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "マクロ"
        
    GoTo ExitHandler
        
End Function



