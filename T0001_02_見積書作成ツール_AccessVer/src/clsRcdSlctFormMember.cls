VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRcdSlctFormMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Class")
Option Compare Database
Option Explicit
Implements clsAbsRcdSlctForm
'**************************
'*レコード選択フォーム処理クラス　人名データ
'**************************

'定数欄
Private Const SOURCE_NAME As String = "clsRcdSlctFormMember"



'変数欄



'******************************************************************************************
'*getter/setter欄
'******************************************************************************************




'******************************************************************************************
'*機能      ：フォームの情報を取得する
'*引数      ：情報格納辞書
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Private Function clsAbsRcdSlctForm_getInfo(pDicInfo As Object) As Boolean

    '定数
    Const FUNC_NAME As String = "clsAbsRcdSlctForm_getInfo"
    
    '変数
    
    On Error GoTo ErrorHandler

    clsAbsRcdSlctForm_getInfo = False
    
    'タイトル
    pDicInfo.Add "formTitle", "担当者検索"
    '案内文
    pDicInfo.Add "announceTxt", "●検索する担当者のレコードを選択してください。"
    'プロパティ　リストボックス
    pDicInfo.Add "rowSourceForList", QRY_QSR01

TruePoint:

    clsAbsRcdSlctForm_getInfo = True

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




'******************************************************************************************
'*機能      ：btnOKClick処理
'*引数      ：対象フォーム
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Private Function clsAbsRcdSlctForm_btnOKClick(ByVal pForm As Object) As Boolean

    '定数
    Const FUNC_NAME As String = "clsAbsRcdSlctForm_btnOKClick"
    
    '変数
    Dim boolErr As Boolean
    
    On Error GoTo ErrorHandler

    clsAbsRcdSlctForm_btnOKClick = False
    
    '空欄チェック
    If Not checkWhetherControlsVacant(boolErr, pForm.lst_selectField.Column(1)) Then GoTo ExitHandler
    If boolErr Then Call MsgBox("リストのレコードを選択してください。", vbExclamation, MESSAGE_TITLE_WARNING): GoTo TruePoint
    
    '呼び出し元に選択値を反映
    Form_F05_見積書項目設定.lst_mmbr.RowSource = "SELECT 人名ID,名前 FROM " & _
        TBL_M_MEMBER & " " & _
        "WHERE 人名ID = " & _
        pForm.lst_selectField.Column(0) & ";"
    
    '画面を閉じる
    DoCmd.Close acForm, pForm.Name, acSaveNo

TruePoint:

    clsAbsRcdSlctForm_btnOKClick = True

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



