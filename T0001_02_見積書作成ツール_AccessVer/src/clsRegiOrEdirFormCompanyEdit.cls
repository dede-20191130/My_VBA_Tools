VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRegiOrEdirFormCompanyEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Class")
Option Compare Database
Option Explicit
Implements clsAbsRegiOrEditForm
'**************************
'*取引先会社編集フォーム処理
'**************************

'定数欄
Private Const SOURCE_NAME As String = "clsRegiOrEdirFormCompanyEdit"

'変数欄
Private objRegiOrEdirFormCommon As New clsRegiOrEdirFormCommon


'******************************************************************************************
'*getter/setter欄
'******************************************************************************************





'******************************************************************************************
'*関数名    ：getInfo_consumptionTax_edit
'*機能      ：フォームの情報を取得する 取引先会社データ 編集
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
    pDicInfo.Add "subFormSourceObject", FormName_04_SUB03
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
'*機能      ：直近の画面の情報を取得する 取引先会社データ　編集
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
        '会社名
        .txt_cmpNm.VALUE = Form_F03_設定.subF_Datasheet.Form.txt_cmpNm.VALUE
        '電話番号
        .txt_telNum.VALUE = Form_F03_設定.subF_Datasheet.Form.txt_telNum.VALUE
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
'*機能      ：編集処理　取引先会社データ
'*引数(1)   ：cancel
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Private Function clsAbsRegiOrEditForm_btnExecuteClick(Cancel As Boolean) As Boolean

    '定数
    Const FUNC_NAME As String = "clsAbsRegiOrEditForm_btnExecuteClick"
    
    '変数
    Dim boolError As Boolean
    Dim cmpNm As String
    Dim telNum As String
    Dim wrs As New clsWrappedRecordSet
    
    On Error GoTo ErrorHandler

    clsAbsRegiOrEditForm_btnExecuteClick = False
    Cancel = False
    
    '確認
    If Not objRegiOrEdirFormCommon.showConfirmation(EDIT, Cancel) Then GoTo ExitHandler
    If Cancel Then: GoTo TruePoint
    
    
    With Form_F04_登録_編集_SUB03_基礎データ_取引先会社データ
        '空欄チェック
        If Not checkWhetherControlsVacant(boolError, .txt_cmpNm.VALUE, .txt_telNum.VALUE) Then GoTo ExitHandler
        If boolError Then Call MsgBox(MESSAGE_EXIST_BLANK, vbExclamation, MESSAGE_TITLE_WARNING): Cancel = True: GoTo TruePoint

        '型チェックなし

        '値チェック
        If Not checkTelNum(.txt_telNum.VALUE, boolError) Then GoTo ExitHandler
        If boolError Then Call MsgBox(objRegiOrEdirFormCommon.MESSAGE_ERR_NOT_TELNUM, vbExclamation, MESSAGE_TITLE_WARNING): Cancel = True: GoTo TruePoint

        cmpNm = .txt_cmpNm.VALUE
        telNum = .txt_telNum.VALUE

    End With
        
    'データの更新
    Set wrs.varRecordset = gObjDB.Database.OpenRecordset("SELECT * FROM" & myVBVacant & TBL_M_ORG & myVBVacant & "WHERE 会社ID =" & myVBVacant & Form_F04_登録_編集.currentID)
    With wrs.varRecordset
        If .EOF Then: GoTo TruePoint
        .EDIT
        .Fields("会社名").VALUE = cmpNm
        .Fields("電話番号").VALUE = telNum
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


