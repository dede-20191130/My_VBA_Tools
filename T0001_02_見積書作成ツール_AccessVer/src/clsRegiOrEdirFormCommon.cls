VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRegiOrEdirFormCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Class")
Option Compare Database
Option Explicit

'**************************
'*登録or編集Form処理クラス共通
'**************************

'定数欄
Private Const SOURCE_NAME As String = "clsRegiOrEdirFormCommon"

'変数欄


'******************************************************************************************
'*getter/setter欄
'******************************************************************************************

Public Property Get MESSAGE_ERR_NOT_TELNUM() As String
    MESSAGE_ERR_NOT_TELNUM = "には数字とハイフンのみを入力してください。"
End Property

Public Property Get MESSAGE_ERR_INTEGRAL() As String
    MESSAGE_ERR_INTEGRAL = "には2,000,000,000以下の数値を入力してください。"
End Property

'******************************************************************************************
'*関数名    ：showConfirmation
'*機能      ：確認メッセージ表示
'*引数(1)   ：cancel
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function showConfirmation( _
    ByVal regiOrEdit As String, _
    ByRef Cancel As Boolean _
) As Boolean
    
    '定数
    Const FUNC_NAME As String = "showConfirmation"
    
    '変数
    
    On Error GoTo ErrorHandler

    showConfirmation = False
    
    'はいを選択しない場合はキャンセル
    If MsgBox( _
        "上記内容で" & _
            regiOrEdit & _
            "を行います。" & _
            vbNewLine & _
            "よろしいですか。", _
        vbYesNo, _
        TOOL_NAME _
    ) <> vbYes Then
        Cancel = True
    End If
        
    showConfirmation = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & err.Number & vbNewLine & _
           err.description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Function

