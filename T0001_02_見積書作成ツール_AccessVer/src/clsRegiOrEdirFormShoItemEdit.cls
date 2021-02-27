VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRegiOrEdirFormShoItemEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Class")
Option Compare Database
Option Explicit
Implements clsAbsRegiOrEditForm
'**************************
'*商品データ編集フォーム処理
'**************************

'定数欄
Private Const SOURCE_NAME As String = "clsRegiOrEdirFormShoItemEdit"

'変数欄
Private objRegiOrEdirFormCommon As New clsRegiOrEdirFormCommon


'******************************************************************************************
'*getter/setter欄
'******************************************************************************************





'******************************************************************************************
'*関数名    ：getInfo_consumptionTax_edit
'*機能      ：フォームの情報を取得する 商品データ 編集
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
    pDicInfo.Add "formTitle", "取引先会社編集"
    'サブフォーム
    pDicInfo.Add "subFormSourceObject", FormName_04_SUB05
    '説明ラベルテキスト
    pDicInfo.Add "announceLblText", "指定した取引先会社の情報を編集してください。"
    '実行ボタンタイトル
    pDicInfo.Add "executeBtnTitle", EDIT

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


'******************************************************************************************
'*関数名    ：clsAbsRegiOrEditForm_getPreviousFormData
'*機能      ：直近の画面の情報を取得する 商品データ　編集
'*引数(1)   ：情報格納Dic
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Private Function clsAbsRegiOrEditForm_getPreviousFormData(currentID As Long) As Boolean

    '定数
    Const FUNC_NAME As String = "clsAbsRegiOrEditForm_getPreviousFormData"
    
    '変数
    
    On Error GoTo ErrorHandler

    clsAbsRegiOrEditForm_getPreviousFormData = False
    
    'ID
    currentID = Form_F03_設定.subF_Datasheet.Form.txt_ID.VALUE
    With Form_F04_登録_編集.subF_cutForm.Form
        '商品コード
        .txt_itemCode.VALUE = Form_F03_設定.subF_Datasheet.Form.txt_itemCode.VALUE
        '商品名
        .txt_itemNm.VALUE = Form_F03_設定.subF_Datasheet.Form.txt_itemNm.VALUE
        '単価
        .txt_itemPrc.VALUE = Form_F03_設定.subF_Datasheet.Form.txt_itemPrc.VALUE
    End With
    clsAbsRegiOrEditForm_getPreviousFormData = True
    
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
'*関数名    ：clsAbsRegiOrEditForm_btnExecuteClick
'*機能      ：編集処理　商品データ
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
        
    'データの更新
    Set wrs.varRecordset = gObjDB.Database.OpenRecordset("SELECT * FROM" & myVBVacant & TBL_M_PROD & myVBVacant & "WHERE 商品ID =" & myVBVacant & Form_F04_登録_編集.currentID)
    With wrs.varRecordset
        If .EOF Then: GoTo TruePoint
        .EDIT
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




