VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableCreater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Class")
Option Explicit

'**************************
'*TableCreater
'**************************


'定数
Private Const HEADER_COLOR = 15917529            'ヘッダーセルの色

'変数
Private myRange As Range                         '範囲
Private myColumnSubTotal As Long                 '小計列番号


'******************************************************************************************
'*getter/setter
'******************************************************************************************


Public Property Set Range(ByVal pRng As Range)
    Set myRange = pRng
End Property


Public Property Let ColumnSubTotal(ByVal num As Long)
    '範囲変数が設定されていない状態で呼出は禁止
    If myRange Is Nothing Then Err.Raise 1000, , "範囲が設定されていません。"
    '引数が範囲オブジェクトの列番号より外の列の番号であればエラー
    If num < myRange.EntireColumn(1).Column Or myRange.EntireColumn(myRange.EntireColumn.Count).Column < num Then Err.Raise 1001, , "列番号の指定が不正です。"
    '値を設定
    myColumnSubTotal = num
End Property


'******************************************************************************************
'*関数名    ：Class_Initialize
'*機能      ：
'*引数      ：
'******************************************************************************************
Private Sub Class_Initialize()
    
    '定数
    Const FUNC_NAME As String = "Class_Initialize"
    
    '変数
    
    On Error GoTo ErrorHandler
    
    '本クラスでは特にやることはない。

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "TableCreater"
        
    GoTo ExitHandler
        
End Sub


'******************************************************************************************
'*関数名    ：drawLines
'*機能      ：罫線を引く
'*引数      ：
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function drawLines() As Boolean
    
    '定数
    Const FUNC_NAME As String = "drawLines"
    
    '変数
    
    On Error GoTo ErrorHandler

    drawLines = False
    
    '範囲変数が設定されていない状態で呼出は禁止
    If myRange Is Nothing Then Err.Raise 1000, , "範囲が設定されていません。"
    
    '罫線を引く
    myRange.Borders.LineStyle = xlContinuous

    drawLines = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "TableCreater"
        
    GoTo ExitHandler
        
End Function


'******************************************************************************************
'*関数名    ：setStyleForHeader
'*機能      ：ヘッダーの強調のためのスタイル変更を行う
'               ヘッダーは与えられた範囲の一行目のセルとする
'*引数      ：
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function setStyleForHeader() As Boolean
    
    '定数
    Const FUNC_NAME As String = "setStyleForHeader"
    
    '変数
    
    On Error GoTo ErrorHandler

    setStyleForHeader = False
    
    '範囲変数が設定されていない状態で呼出は禁止
    If myRange Is Nothing Then Err.Raise 1000, , "範囲が設定されていません。"
    
    'スタイル変更
    With myRange.Rows(1)
        '範囲の一行目のセルの色を変更
        .Interior.color = HEADER_COLOR
        '範囲の一行目を太字にする
        .Font.Bold = True
        '範囲の一行目を中央揃えにする
        .HorizontalAlignment = xlCenter
    End With
        
    
    setStyleForHeader = True
    
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "TableCreater"
        
    GoTo ExitHandler
        
End Function




'******************************************************************************************
'*関数名    ：calcTotalFromSubTotal
'*機能      ：小計列から合計を計算し出力する
'*引数      ：
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function calcTotalFromSubTotal() As Boolean
    
    '定数
    Const FUNC_NAME As String = "calcTotalFromSubTotal"
    
    '変数
    Dim sumVal As Long
    Dim cell As Range
    Dim subTotalOrder As Long
    
    On Error GoTo ErrorHandler

    calcTotalFromSubTotal = False
    
    '範囲変数が設定されていない状態で呼出は禁止
    If myRange Is Nothing Then Err.Raise 1000, , "範囲が設定されていません。"
    '小計列が設定されていない状態で呼出は禁止
    If myColumnSubTotal = 0 Then Err.Raise 1002, , "小計列が設定されていません。"
        
    '小計列が範囲の列の何番目であるかを算出
    subTotalOrder = myColumnSubTotal - myRange(1).Column + 1
        
    '合計値計算（ヘッダー列は含まない）
    For Each cell In myRange.Columns(subTotalOrder).Cells.Offset(1).Resize(myRange.Columns(subTotalOrder).Cells.Offset(1).Cells.Count - 1)
        '数値の場合のみ加算
        If IsNumeric(cell.Value) Then sumVal = sumVal + cell.Value
    Next cell
    If sumVal = 0 Then GoTo TruePoint
    
    '合計値を小計列の最下端セルの下のセルに記入
    With myRange.Columns(subTotalOrder).Rows(myRange.Columns(subTotalOrder).Cells.Count).Offset(1)
        .Value = sumVal
        '合計ラベルセル
        With .Offset(, -1)
            'ラベルを記入
            .Value = "合計"
            '合計ラベルセルと合計セルに罫線を引く
            .Resize(, .Columns.Count + 1).Borders.LineStyle = xlContinuous
        End With
        
    End With
       
TruePoint:
       
    calcTotalFromSubTotal = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "TableCreater"
        
    GoTo ExitHandler
        
End Function


