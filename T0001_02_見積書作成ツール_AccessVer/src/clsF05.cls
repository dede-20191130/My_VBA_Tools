VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsF05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Class")
Option Compare Database
Option Explicit

'**************************
'*見積書項目設定Form処理クラス
'**************************

'定数欄
Private Const SOURCE_NAME As String = "clsF05"

'変数欄
Private objCmbEvent As clsCmbListener


'******************************************************************************************
'*getter/setter欄
'******************************************************************************************



'******************************************************************************************
'*機能      ：Form_Load処理
'*引数      ：フォーム
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function formLoad(ByVal pForm As Access.Form) As Boolean
    
    '定数
    Const FUNC_NAME As String = "formLoad"
    
    '変数
    
    On Error GoTo ErrorHandler

    formLoad = False
    
    'サイズ
    Call DoCmd.MoveSize(, , , 567 * 24)
    
    'イベントリスナー
    Set objCmbEvent = New clsCmbListener: Set objCmbEvent.ComboBox = pForm.cmb_csmpTax


TruePoint:

    formLoad = True

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
'*機能      ：btnShowPrevwClick処理
'*引数      ：Database
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function btnShowPrevwClick(ByVal daoDB As Database) As Boolean
    
    '定数
    Const FUNC_NAME As String = "btnShowPrevwClick"
    
    '変数
    Dim filePath As String
    
    On Error GoTo ErrorHandler

    btnShowPrevwClick = False
    
    filePath = CurrentProject.Path & _
        "\" & _
        "template.xlsx"
    
    If Not mdlFile.saveBookEstmTmpl(daoDB, filePath) Then GoTo ExitHandler
    If Not Me.openBookReadOnly(filePath) Then GoTo ExitHandler

TruePoint:

    btnShowPrevwClick = True

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
'*機能      ：btnNextClick処理
'*引数      ：自フォーム
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function btnNextClick(ByVal pForm As Access.Form) As Boolean
    
    '定数
    Const FUNC_NAME As String = "btnNextClick"
    
    '変数
    Dim boolErr As Boolean
    Dim objF05DtTrsfr As New clsF05DtTrsfr
    
    On Error GoTo ErrorHandler

    btnNextClick = False
    
    
    With pForm
        'チェック処理
        If Not checkWhetherControlsVacant( _
            boolErr, .txt_estiNum.VALUE, .lst_cmpNm.Column(0, 0), .lst_mmbr.Column(0, 0), _
            .txt_delivery.VALUE, .txt_payCondition.VALUE, .txt_validLimit.VALUE, _
            .cmb_csmpTax.VALUE _
        ) Then GoTo ExitHandler
        If boolErr Then Call MsgBox(MESSAGE_EXIST_BLANK, vbExclamation, MESSAGE_TITLE_WARNING): GoTo TruePoint
    End With
    
    'データ転送オブジェクトにデータ格納
    objF05DtTrsfr.estiNum = pForm.txt_estiNum.VALUE
    objF05DtTrsfr.cmpNm = pForm.lst_cmpNm.Column(0, 0)
    objF05DtTrsfr.mmbr = pForm.lst_mmbr.Column(0, 0)
    objF05DtTrsfr.delivery = pForm.txt_delivery.VALUE
    objF05DtTrsfr.payCondition = pForm.txt_payCondition.VALUE
    objF05DtTrsfr.validLimit = pForm.txt_validLimit.VALUE
    objF05DtTrsfr.other = Nz(pForm.txt_other.VALUE, "")
    objF05DtTrsfr.csmpTax = pForm.cmb_csmpTax.VALUE
    
    '次画面がロード済みの場合
    If Application.CurrentProject.AllForms(FormName_08).IsLoaded Then
        '作成済みのオブジェクトを置換する
        Set Form_F08_見積書_表作成.gF08ObjF08.gObjF05DtTrsfr = objF05DtTrsfr
        'フォームの情報設定
        Form_F08_見積書_表作成.txt_estiNum.VALUE = objF05DtTrsfr.estiNum
    'それ以外の場合
    Else
        'データ転送オブジェクトをグローバルなスコープでキープ
        If Not gObjDtTrsfrManager.setObjDtTrsfr(objF05DtTrsfr) Then GoTo ExitHandler
    End If
    
    '画面遷移
    Forms(pForm.Name).Visible = False
    If Not showFormInvisibleOrUnloaded(FormName_08, pForm.Name) Then GoTo ExitHandler
    
    
TruePoint:

    btnNextClick = True

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
'*機能      ：btnBackClick処理
'*引数      ：自フォーム
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function btnBackClick(ByVal pForm As Access.Form, ByVal daoDB As dao.Database) As Boolean
    
    '定数
    Const FUNC_NAME As String = "btnBackClick"
    
    '変数
    
    On Error GoTo ErrorHandler

    btnBackClick = False
    
    '一時テーブル削除
    Call daoDB.execute("DELETE FROM " & TBL_W_ESTM & ";")
    
    '関連画面をすべて閉じる
    If Not closeFormIfLoaded(pForm.Name, FormName_08) Then GoTo ExitHandler
    
    'メニューを表示する
    Form_F02_メニュー.Visible = True


TruePoint:

    btnBackClick = True

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
'*機能      ：見積書テンプレートを読み取り専用で開く
'*引数      ：ファイルパス
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function openBookReadOnly(ByVal filePath As String)
    '定数
    Const FUNC_NAME As String = "openBookReadOnly"
    
    '変数
    
    On Error GoTo ErrorHandler

    openBookReadOnly = False
    
    '読み取り専用で開く
    Dim myXlApp As Object: Set myXlApp = CreateObject("Excel.Application")
    myXlApp.Visible = True
    Call myXlApp.Workbooks.Open(filePath, , True)
    
    '一時保存ブック削除
    Kill filePath

TruePoint:

    openBookReadOnly = True

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

