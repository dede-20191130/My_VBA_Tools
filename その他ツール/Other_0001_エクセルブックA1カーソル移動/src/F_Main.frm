VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_Main 
   Caption         =   "Other_0001_エクセルブックA1カーソル移動"
   ClientHeight    =   4920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9780.001
   OleObjectBlob   =   "F_Main.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "F_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Explicit


'******************************************************************************************
'*関数名    ：UserForm_Terminate
'*機能      ：フォームを閉じる際にツールを終了する
'*引数(1)   ：
'******************************************************************************************
Private Sub UserForm_Terminate()

    
    '定数
    Const FUNC_NAME As String = "UserForm_Terminate"
    
    '変数
    
    On Error GoTo ErrorHandler
    
    #If Not CBool(DEBUG_MODE) Then
        ThisWorkbook.Close False
    #End If
    
ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Sub


'******************************************************************************************
'*関数名    ：btn_execute_Click
'*機能      ：
'*引数(1)   ：
'******************************************************************************************
Private Sub btn_execute_Click()
    
    '定数
    Const FUNC_NAME As String = "btn_execute_Click"
    
    '変数
    Dim suffix As String
    Dim folderPath As String
    Dim objFormatExcel As clsFormatExcel
    
    On Error GoTo ErrorHandler
    
    Set objFormatExcel = New clsFormatExcel
    
    'フォルダ選択ダイアログ
    folderPath = getFolderPathFromDialog("フォルダを指定してください")
    If folderPath = "" Then GoTo ExitHandler
        
    '再帰的なファイル検索の有無
    suffix = _
           WorksheetFunction.Rept(Me.opt_onRecurse.Tag, Abs(CLng(CBool(Me.opt_onRecurse.Value)))) & _
           WorksheetFunction.Rept(Me.opt_offRecurse.Tag, Abs(CLng(CBool(Me.opt_offRecurse.Value))))
    If suffix = "" Then GoTo ExitHandler
        
    '処理関数の呼出
    If Not CallByName(objFormatExcel, FUNC_NAME & "_" & suffix, VbMethod, folderPath) Then GoTo ExitHandler
        
    '処理件数の通知
    MsgBox objFormatExcel.tgtFileCnt & "件のExcelファイルをカーソル移動しました。", , TOOL_NAME

ExitHandler:
    
    Set objFormatExcel = Nothing
    
    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Sub


'******************************************************************************************
'*関数名    ：btn_terminate_Click
'*機能      ：
'*引数(1)   ：
'******************************************************************************************
Private Sub btn_terminate_Click()
    
    '定数
    Const FUNC_NAME As String = "btn_terminate_Click"
    
    '変数
    
    On Error GoTo ErrorHandler
    
    'メッセージ
    If MsgBox("ツールを終了します。", vbYesNo, TOOL_NAME) <> vbYes Then GoTo ExitHandler
    
    'フォームを閉じる
    Unload Me
    

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Sub




