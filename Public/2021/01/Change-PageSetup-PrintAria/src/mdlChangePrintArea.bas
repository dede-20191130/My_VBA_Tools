Attribute VB_Name = "mdlChangePrintArea"
Option Explicit

'**************************
'*PageSetup_PrintArea変更テスト
'**************************

'定数欄
Private Const SOURCE_NAME As String = "mdlChangePrintArea"


'変数欄



'******************************************************************************************
'*getter/setter欄
'******************************************************************************************





'******************************************************************************************
'*関数名    ：changePrintAreaBeforeRevised
'*機能      ：PrintAreaをひとつ下の行に変更する 修正前
'*引数      ：
'******************************************************************************************
Public Sub changePrintAreaBeforeRevised()
    
    '定数
    Const FUNC_NAME As String = "changePrintAreaBeforeRevised"
    
    '変数
    Dim prePrintAreaAddress As String
    
    On Error GoTo ErrorHandler
    
    With ThisWorkbook.Worksheets(1)
    
        '現在の印刷範囲アドレス
        prePrintAreaAddress = .PageSetup.PrintArea
        
        '印刷範囲をひとつ下の行に変更する
        .PageSetup.PrintArea = .Range(prePrintAreaAddress).Resize(.Range(prePrintAreaAddress).Rows.Count + 1).Address
        
        Debug.Print .PageSetup.PrintArea
        
    End With

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, SOURCE_NAME
        
    GoTo ExitHandler
        
End Sub





'******************************************************************************************
'*関数名    ：changePrintAreaBeforeRevised
'*機能      ：PrintAreaをひとつ下の行に変更する 修正01
'*引数      ：
'******************************************************************************************
Public Sub changePrintAreaRevised01()
    
    '定数
    Const FUNC_NAME As String = "changePrintAreaRevised01"
    
    '変数
    Dim prePrintAreaAddress As String
    Dim currentStyle As XlReferenceStyle

    On Error GoTo ErrorHandler
    
    With ThisWorkbook.Worksheets(1)
        
        'セルの参照形式をA1形式に変更
        currentStyle = Application.ReferenceStyle
        Application.ReferenceStyle = xlA1
        
        '現在の印刷範囲アドレス
        prePrintAreaAddress = .PageSetup.PrintArea
        
        '印刷範囲をひとつ下の行に変更する
        .PageSetup.PrintArea = .Range(prePrintAreaAddress).Resize(.Range(prePrintAreaAddress).Rows.Count + 1).Address
        
        Debug.Print .PageSetup.PrintArea
        
        'セルの参照形式を復旧する
        Application.ReferenceStyle = currentStyle
        
    End With

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, SOURCE_NAME
        
    GoTo ExitHandler
        
End Sub




'******************************************************************************************
'*関数名    ：changePrintAreaBeforeRevised
'*機能      ：PrintAreaをひとつ下の行に変更する 修正02
'*引数      ：
'******************************************************************************************
Public Sub changePrintAreaRevised02()
    
    '定数
    Const FUNC_NAME As String = "changePrintAreaRevised02"
    
    '変数
    Dim prePrintAreaAddress As String
    
    On Error GoTo ErrorHandler
    
    With ThisWorkbook.Worksheets(1)
    
        '現在の印刷範囲アドレス
        prePrintAreaAddress = .PageSetup.PrintArea
        
        'アドレスをxlA1参照形式のものに修正
        If Application.ReferenceStyle = xlR1C1 Then prePrintAreaAddress = Application.ConvertFormula(prePrintAreaAddress, xlR1C1, xlA1)
        
        '印刷範囲をひとつ下の行に変更する
        .PageSetup.PrintArea = .Range(prePrintAreaAddress).Resize(.Range(prePrintAreaAddress).Rows.Count + 1).Address
        
        Debug.Print .PageSetup.PrintArea
        
    End With

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, SOURCE_NAME
        
    GoTo ExitHandler
        
End Sub

