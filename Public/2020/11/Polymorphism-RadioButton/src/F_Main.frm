VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_Main 
   Caption         =   "F_Main"
   ClientHeight    =   5220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9330.001
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
'*関数名    ：btn_execute_Click
'*機能      ：
'*引数(1)   ：
'******************************************************************************************
Private Sub btn_execute_Click()
    
    '定数
    Const FUNC_NAME As String = "btn_execute_Click"
    
    '変数
    Dim suffix As String
    Dim objPolymo As clsPolymo
    
    On Error GoTo ErrorHandler
    
    Set objPolymo = New clsPolymo
    
    '選択された処理を取得
    '再帰的なファイル検索の有無
    suffix = _
           WorksheetFunction.Rept(Me.rdo_showCurrent.Tag, Abs(CLng(CBool(Me.rdo_showCurrent.Value)))) & _
           WorksheetFunction.Rept(Me.rdo_showUser.Tag, Abs(CLng(CBool(Me.rdo_showUser.Value)))) & _
           WorksheetFunction.Rept(Me.rdo_showGreeting.Tag, Abs(CLng(CBool(Me.rdo_showGreeting.Value))))
    If suffix = "" Then MsgBox "ラジオボタンの選択が不正です", vbCritical, Tool_Name: GoTo ExitHandler
    
    '処理関数の呼出
    If Not CallByName(objPolymo, FUNC_NAME & "_" & suffix, VbMethod) Then GoTo ExitHandler
    

ExitHandler:
    
    Set objPolymo = Nothing
    
    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, Tool_Name
        
    GoTo ExitHandler
        
End Sub

