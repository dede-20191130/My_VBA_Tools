VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRegiOrEdirFormShoItemRegi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Class")
Option Compare Database
Option Explicit
Implements clsAbsRegiOrEditForm
'**************************
'*商品登録フォーム処理
'**************************

'定数欄
Private Const SOURCE_NAME As String = "clsRegiOrEdirFormShoItemRegi"

'変数欄
Private objRegiOrEdirFormCommon As New clsRegiOrEdirFormCommon


'******************************************************************************************
'*getter/setter欄
'******************************************************************************************






'******************************************************************************************
'*関数名    ：getInfo_numUnit_register
'*機能      ：フォームの情報を取得する 商品 登録
'*引数(1)   ：情報格納Dic
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Private Function clsAbsRegiOrEditForm_getInfo(pDicInfo As Object) As Boolean

    '定数
    Const FUNC_NAME As String = "clsAbsRegiOrEditForm_getInfo"
    
    '変数
    
    On Error GoTo ErrorHandler

    clsAbsRegiOrEditForm_getInfo = False
    
    'タイトル
    pDicInfo.Add "formTitle", "商品データ登録"
    'サブフォーム
    pDicInfo.Add "subFormSourceObject", FormName_04_SUB05
    '説明ラベルテキスト
    pDicInfo.Add "announceLblText", "新しい商品データを入力してください。"
    '実行ボタンタイトル
    pDicInfo.Add "executeBtnTitle", REGISTER

    clsAbsRegiOrEditForm_getInfo = True
    
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



Private Function clsAbsRegiOrEditForm_getPreviousFormData(currentID As Long) As Boolean
    'nothing to do
    clsAbsRegiOrEditForm_getPreviousFormData = True
End Function




'******************************************************************************************
'*関数名    ：clsAbsRegiOrEditForm_btnExecuteClick
'*機能      ：登録処理　商品
'*引数(1)   ：cancel
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Private Function clsAbsRegiOrEditForm_btnExecuteClick(Cancel As Boolean) As Boolean

    '定数
    Const FUNC_NAME As String = "clsAbsRegiOrEditForm_btnExecuteClick"
    
    '変数
    Dim boolError As Boolean
    Dim itemCode As String
    Dim itemNm As String
    Dim itemPrc As Long
    Dim wrs As New clsWrappedRecordSet
    
    On Error GoTo ErrorHandler

    clsAbsRegiOrEditForm_btnExecuteClick = False
    Cancel = False
    
    '確認
    If Not objRegiOrEdirFormCommon.showConfirmation(REGISTER, Cancel) Then GoTo ExitHandler
    If Cancel Then: GoTo TruePoint
    
    
    With Form_F04_登録_編集_SUB05_基礎データ_商品データ
        '空欄チェック
        If Not checkWhetherControlsVacant(boolError, .txt_itemCode.VALUE, .txt_itemNm.VALUE, .txt_itemPrc.VALUE) Then GoTo ExitHandler
        If boolError Then Call MsgBox(MESSAGE_EXIST_BLANK, vbExclamation, MESSAGE_TITLE_WARNING): Cancel = True: GoTo TruePoint

        '型チェック
        If Not checkType(.txt_itemPrc.VALUE, dbLong, boolError) Then GoTo ExitHandler
        If boolError Then Call MsgBox(.txt_itemPrc.Tag & objRegiOrEdirFormCommon.MESSAGE_ERR_INTEGRAL, vbExclamation, MESSAGE_TITLE_WARNING): Cancel = True: GoTo TruePoint

        '値チェックなし

        itemCode = .txt_itemCode.VALUE
        itemNm = .txt_itemNm.VALUE
        itemPrc = .txt_itemPrc.VALUE

    End With
        
    'データの挿入
    Set wrs.varRecordset = gObjDB.Database.OpenRecordset(TBL_M_PROD)
    With wrs.varRecordset
        .AddNew
        .Fields("商品コード").VALUE = itemCode
        .Fields("商品名").VALUE = itemNm
        .Fields("単価").VALUE = itemPrc
        .Update
    End With
    
    Forms(Form_F04_登録_編集.gPreviousScreenName).Refresh

TruePoint:

    clsAbsRegiOrEditForm_btnExecuteClick = True
    
ExitHandler:
    
    Exit Function
    
ErrorHandler:

    Call mdlMessage.showErrMessageUpdateRcd(err.Number, err.description)
        
    GoTo ExitHandler
        
End Function



