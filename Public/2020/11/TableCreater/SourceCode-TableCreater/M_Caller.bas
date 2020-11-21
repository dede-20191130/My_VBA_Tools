Attribute VB_Name = "M_Caller"
'@Folder("Module")
Option Explicit

'**************************
'*TableCreaterクラスの呼び出し元
'**************************

'定数

'変数



'******************************************************************************************
'*getter/setter
'******************************************************************************************


'******************************************************************************************
'*関数名    ：TestTemplateA
'*機能      ：原本のテンプレAについて、TableCreaterを用いて表を作成する
'               作成場所：新規シート
'*引数      ：
'******************************************************************************************
Public Sub TestTemplateA()
    
    '定数
    Const FUNC_NAME As String = "TestTemplateA"
    
    '変数
    Dim ws As Worksheet
    Dim objTableCreater As tableCreater
    
    On Error GoTo ErrorHandler
    
    With ThisWorkbook
        '新規シート作成
        Set ws = .Worksheets.Add(, .Worksheets(.Worksheets.Count))
        ws.Name = "テンプレA_" & Format(Now, "yyyymmddhhnnss")
        
        '原本よりテンプレをコピー
        ws.Range(ws.Cells(2, 2), ws.Cells(9, 4)).Value = .Worksheets("原本").Range(.Worksheets("原本").Cells(2, 2), .Worksheets("原本").Cells(9, 4)).Value
        
        'TableCreaterをオブジェクト化
        Set objTableCreater = New tableCreater
        '範囲と小計列を設定
        Set objTableCreater.Range = ws.Range(ws.Cells(2, 2), ws.Cells(9, 4))
        objTableCreater.ColumnSubTotal = 4
        
        '罫線を引く 異常終了時はExitHandler（終了処理）に移行
        If Not objTableCreater.drawLines Then GoTo ExitHandler
         
        'ヘッダーの強調のためのスタイル変更を行う 異常終了時はExitHandler（終了処理）に移行
        If Not objTableCreater.setStyleForHeader Then GoTo ExitHandler
        
        '小計から合計を計算 異常終了時はExitHandler（終了処理）に移行
        If Not objTableCreater.calcTotalFromSubTotal Then GoTo ExitHandler
        
        '列幅の調整
        ws.Range(ws.Cells(2, 2), ws.Cells(9, 4)).EntireColumn.AutoFit
        
    End With
    

ExitHandler:
    
    '変数を解放
    Set objTableCreater = Nothing
    Set ws = Nothing
    
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
'*関数名    ：TestTemplateB
'*機能      ：原本のテンプレBについて、TableCreaterを用いて表を作成する
'               作成場所：新規シート
'*引数      ：
'******************************************************************************************
Public Sub TestTemplateB()
    
    '定数
    Const FUNC_NAME As String = "TestTemplateB"
    
    '変数
    Dim ws As Worksheet
    Dim objTableCreater As tableCreater
    
    On Error GoTo ErrorHandler
    
    With ThisWorkbook
        '新規シート作成
        Set ws = .Worksheets.Add(, .Worksheets(.Worksheets.Count))
        ws.Name = "テンプレB_" & Format(Now, "yyyymmddhhnnss")
        
        '原本よりテンプレをコピー
        ws.Range(ws.Cells(2, 2), ws.Cells(8, 10)).Value = .Worksheets("原本").Range(.Worksheets("原本").Cells(12, 2), .Worksheets("原本").Cells(18, 10)).Value
        
        'TableCreaterをオブジェクト化
        Set objTableCreater = New tableCreater
        '範囲と小計列を設定
        Set objTableCreater.Range = ws.Range(ws.Cells(2, 2), ws.Cells(8, 10))
        objTableCreater.ColumnSubTotal = 10
        
        '罫線を引く 異常終了時はExitHandler（終了処理）に移行
        If Not objTableCreater.drawLines Then GoTo ExitHandler
         
        'ヘッダーの強調のためのスタイル変更を行う 異常終了時はExitHandler（終了処理）に移行
        If Not objTableCreater.setStyleForHeader Then GoTo ExitHandler
        
        '小計から合計を計算 異常終了時はExitHandler（終了処理）に移行
        If Not objTableCreater.calcTotalFromSubTotal Then GoTo ExitHandler
        
        '列幅の調整
        ws.Range(ws.Cells(2, 2), ws.Cells(8, 10)).EntireColumn.AutoFit
        
    End With
    

ExitHandler:
    
    '変数を解放
    Set objTableCreater = Nothing
    Set ws = Nothing
    
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

