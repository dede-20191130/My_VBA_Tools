VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_Create_Estimation_Docs 
   Caption         =   "見積書作成"
   ClientHeight    =   3435
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7560
   OleObjectBlob   =   "F_Create_Estimation_Docs.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "F_Create_Estimation_Docs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Explicit

'******************************************************************************************
'*関数名    ：CommandButton_Execute_Create_Click
'*機能      ：〇〇○
'*引数(1)   ：〇〇○
'******************************************************************************************
Private Sub CommandButton_Execute_Create_Click()

    '定数
    Const FUNC_NAME As String = "CommandButton_Execute_Create_Click"
    
    '変数
    Dim err_str As String
    
    On Error GoTo ErrorHandler
    '---以下に処理を記述---
    
    '高速化
    If Not Execute_SpeedUp() Then GoTo ExitHandler
    
    '作成前バリデーションチェック
    err_str = Is_Valid_Main
    If err_str <> "" Then
        MsgBox ERR_MSG_CREATED_DOCS_MAIN_HEDD & err_str, vbExclamation, TOOL_NAME
        GoTo ExitHandler
    End If
    
    '見積書作成
    err_str = Create_Estimate_Docs_Main
    If err_str <> "" Then
        MsgBox ERR_MSG_CREATED_DOCS_MAIN_EACH_HEDD & vbLf & vbLf & err_str, vbExclamation, TOOL_NAME
    End If
    
    Unload F_Create_Estimation_Docs
    
ExitHandler:
    
    '復旧
    Call Reset_SpeedUp
    
    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生しましたのでマクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号" & Err.Number & Chr(13) & Err.Description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Sub

'******************************************************************************************
'*関数名    ：UserForm_Initialize
'*機能      ：初期化時実行
'*引数(1)   ：
'******************************************************************************************
Private Sub UserForm_Initialize()
    
    '定数
    Const FUNC_NAME As String = "UserForm_Initialize"
    
    '変数
    Dim max_data_num As Long
    Dim arr_data_num() As Long
    Dim i As Long
    
    '---以下に処理を記述---
    
    '# コンボボックスにデータ番号を格納
    '## データ番号取得
    max_data_num = Get_Current_Max_Estimate_Data_Num()
    '## 格納
    ReDim arr_data_num(1 To max_data_num)
    For i = 1 To max_data_num
        arr_data_num(i) = i
    Next i
    ComboBox_Target_Num_Start.List = arr_data_num
    ComboBox_Target_Num_End.List = arr_data_num
    
        
End Sub


