VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsF08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Class")
Option Compare Database
Option Explicit

'**************************
'*見積書_表作成Form処理クラス
'**************************

'定数欄
Private Const SOURCE_NAME As String = "clsF08"

'変数欄
Private myObjF05DtTrsfr As clsF05DtTrsfr


'******************************************************************************************
'*getter/setter欄
'******************************************************************************************
Public Property Set gObjF05DtTrsfr(ByVal RHS As clsF05DtTrsfr)
    Set myObjF05DtTrsfr = RHS
End Property



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
    
    'フォームデータの初期化
    Set myObjF05DtTrsfr = gObjDtTrsfrManager.getObjDtTrsfr("clsF05DtTrsfr")
    
    'フォームの情報設定
    pForm.txt_estiNum.VALUE = myObjF05DtTrsfr.estiNum
    pForm.lbl_alertBlank.Visible = True
    

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
'*機能      ：btn_addRcd_Click処理
'*引数      ：フォーム
'*引数      ：データベース
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function btnAddRcdClick(ByVal pForm As Access.Form, ByVal daoDB As dao.Database) As Boolean
    
    '定数
    Const FUNC_NAME As String = "btnAddRcdClick"
    
    '変数
    Dim wrs As New clsWrappedRecordSet
    Dim nextNumber As Long
    
    On Error GoTo ErrorHandler

    btnAddRcdClick = False
    
    'W_見積_表作成用データ　次に採番される番号を取得
    Set wrs.varRecordset = daoDB.OpenRecordset(QRY_Q_MAX_W_ESTM)
    nextNumber = wrs.varRecordset.Fields(0).VALUE + 1
    Set wrs.varRecordset = Nothing
    
    '画面表示
    DoCmd.OpenForm FormName_09, , , , , acDialog, _
        pForm.Name & ";" & nextNumber
    
    '画面更新
    pForm.subF_currentItem.Form.Requery
    
    '警告表示の表示有無
    Set wrs.varRecordset = daoDB.OpenRecordset(QRY_Q_MAX_W_ESTM)
    pForm.lbl_alertBlank.Visible = (wrs.varRecordset.Fields(0).VALUE <= 0)
    
TruePoint:

    btnAddRcdClick = True

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
'*機能      ：btn_deleteRcd_Click処理
'*引数      ：
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function btnDeleteRcdClick(ByVal pForm As Access.Form, ByVal daoDB As dao.Database) As Boolean
    
    '定数
    Const FUNC_NAME As String = "btnDeleteRcdClick"
    
    '変数
    Dim wrs As New clsWrappedRecordSet
    
    On Error GoTo ErrorHandler

    btnDeleteRcdClick = False
    
    'W_見積_表作成用データ　レコード数をチェック
    Call daoDB.TableDefs.Refresh
    If daoDB.TableDefs(TBL_W_ESTM).RecordCount < 1 Then MsgBox "作成済みのレコードが存在しません。", vbExclamation, MESSAGE_TITLE_WARNING: GoTo TruePoint
    
    '画面表示
    DoCmd.OpenForm FormName_11, , , , , acDialog, _
        pForm.Name
    
    '画面更新
    pForm.subF_currentItem.Form.Requery
    
    '警告表示の表示有無
    Set wrs.varRecordset = daoDB.OpenRecordset(QRY_Q_MAX_W_ESTM)
    pForm.lbl_alertBlank.Visible = (wrs.varRecordset.Fields(0).VALUE <= 0)
    

TruePoint:

    btnDeleteRcdClick = True

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
'*機能      ：btn_createFile_Click処理
'*引数      ：
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function btnCreateFileClick(ByVal pForm As Access.Form, ByVal daoDB As dao.Database) As Boolean
    
    '定数
    Const FUNC_NAME As String = "btnCreateFileClick"
    Const msoFileDialogFolderPicker As Integer = 4
    
    '変数
    Dim dstPath As String
    Dim estimSht As Object
    Dim filePath As String
    Dim objManageExcelBook As New clsManageExcelBook
    
    
    On Error GoTo ErrorHandler

    btnCreateFileClick = False
    
    'W_見積_表作成用データ　レコード数をチェック
    Call daoDB.TableDefs.Refresh
    If daoDB.TableDefs(TBL_W_ESTM).RecordCount < 1 Then MsgBox "作成済みの[商品・サービス]リストが存在しません。", vbExclamation, MESSAGE_TITLE_WARNING: GoTo TruePoint
    
    '作成先フォルダを設定
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = Environ("USERPROFILE")
        .AllowMultiSelect = False

        .Title = "作成先フォルダを選択してください。"
        If CBool(.Show) = True Then
            dstPath = .SelectedItems(1)
        Else
            GoTo TruePoint
        End If
    End With
    filePath = dstPath & _
        "\" & _
        "見積書_" & _
        Format(Now, "yyyymmddhhnnss") & _
        ".xlsx"
    
    'テンプレートブックを保存
    If Not mdlFile.saveBookEstmTmpl(daoDB, filePath) Then GoTo ExitHandler
    
    'テンプレートブックを開く
    If objManageExcelBook.addExistingBook(filePath) Is Nothing Then GoTo ExitHandler
    
    'ブックに情報を格納
    If Not Me.loadEstimationTable(objManageExcelBook.WorkSheets(1), daoDB) Then GoTo ExitHandler
    
    '作成した見積書のレコードとしてテーブルを更新
    If Not Me.insertEstmTblDtl(daoDB) Then GoTo ExitHandler
    If Not Me.insertEstmTbl(daoDB) Then GoTo ExitHandler
    
    '作成完了
    MsgBox "見積書が作成されました。"
    
TruePoint:

    btnCreateFileClick = True

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
'*機能      ：btn_backPrev_Click処理
'*引数      ：
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function btnBackPrevClick(ByVal pForm As Access.Form, ByVal previousScreenName As String) As Boolean
    
    '定数
    Const FUNC_NAME As String = "btnBackPrevClick"
    
    '変数
    
    On Error GoTo ErrorHandler

    btnBackPrevClick = False
    
    '自画面を非表示
    pForm.Visible = False
    
    '前画面を開く
    If Not showFormInvisibleOrUnloaded(previousScreenName) Then GoTo ExitHandler

TruePoint:

    btnBackPrevClick = True

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
'*機能      ：btn_backMenu_Click処理
'*引数      ：
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function btnBackMenuClick(ByVal pForm As Access.Form, ByVal daoDB As dao.Database) As Boolean
    
    '定数
    Const FUNC_NAME As String = "btnBackMenuClick"
    
    '変数
    
    On Error GoTo ErrorHandler

    btnBackMenuClick = False
    
    '一時テーブル削除
    Call daoDB.execute("DELETE FROM " & TBL_W_ESTM & ";")
    
    '関連画面をすべて閉じる
    If Not closeFormIfLoaded(pForm.Name, FormName_05) Then GoTo ExitHandler
    
    'メニューを表示する
    Form_F02_メニュー.Visible = True


TruePoint:

    btnBackMenuClick = True

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
'*機能      ：
'*引数      ：
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function loadEstimationTable(ByVal ws As Object, ByVal daoDB As dao.Database) As Boolean
    
    '定数
    Const FUNC_NAME As String = "loadEstimationTable"
    Const ESTIMATE_SHEET_NAME_BASE As String = "見積書_"
    
    '変数
    Dim wrs As New clsWrappedRecordSet
    Dim product_table_1st_row_num As Long
    Dim product_table_1st_column_num As Long
    Dim product_table_item_num As Long
    
    On Error GoTo ErrorHandler

    loadEstimationTable = False
    
    '情報を格納
    With ws
    
        'リネーム
        .Name = ESTIMATE_SHEET_NAME_BASE & myObjF05DtTrsfr.estiNum
        
        'SET 見積日
        .Range(STR_NAME_RANGE_TMPL_ESTIMATE_DATE).VALUE = replace(.Range(STR_NAME_RANGE_TMPL_ESTIMATE_DATE).VALUE, "$estimation_day", Format(Now, "yyyy年mm月dd日"))
        
        'SET 消費税
        .Range(STR_NAME_RANGE_TMPL_TAX).VALUE = DLookup("VALUE", TBL_M_BSC_TAX, "ID = " & myObjF05DtTrsfr.csmpTax)
        
        'SET 見積データ
        .Range(STR_NAME_RANGE_TMPL_NUMBER).VALUE = replace(.Range(STR_NAME_RANGE_TMPL_NUMBER).VALUE, "$estimation_serial_num", myObjF05DtTrsfr.estiNum)
        .Range(STR_NAME_RANGE_TMPL_COMPANY).VALUE = replace(.Range(STR_NAME_RANGE_TMPL_COMPANY).VALUE, "$company_name", DLookup("会社名", TBL_M_ORG, "会社ID = " & myObjF05DtTrsfr.cmpNm))
        .Range(STR_NAME_RANGE_TMPL_PERSON_IN_CHARGE).VALUE = replace(.Range(STR_NAME_RANGE_TMPL_PERSON_IN_CHARGE).VALUE, "$parson_in_charge", DLookup("名前", TBL_M_MEMBER, "人名ID = " & myObjF05DtTrsfr.mmbr))
        .Range(STR_NAME_RANGE_TMPL_DELIVERY_DATE).VALUE = myObjF05DtTrsfr.delivery
        .Range(STR_NAME_RANGE_TMPL_PAYMENT_TERMS).VALUE = myObjF05DtTrsfr.payCondition
        .Range(STR_NAME_RANGE_TMPL_EXPIRATION_DATE).VALUE = myObjF05DtTrsfr.validLimit
        .Range(STR_NAME_RANGE_TMPL_OTHER_NOTES).VALUE = myObjF05DtTrsfr.other

        'テーブルのプロパティ
        With .Range(STR_NAME_RANGE_PRODUCT_TABLE_HEDD)
            product_table_item_num = .Columns.Count
            product_table_1st_row_num = .Item(1).Row
            product_table_1st_column_num = .Item(1).Column
        End With

        'SET 見積商品セットデータ
        Dim i As Long
        Dim cnt_num As Long
        cnt_num = 1
        Set wrs.varRecordset = daoDB.OpenRecordset("SELECT * FROM " & TBL_W_ESTM & " ORDER BY Number")
        wrs.varRecordset.MoveFirst
        Do Until wrs.varRecordset.EOF
            .Cells(product_table_1st_row_num + cnt_num, product_table_1st_column_num).VALUE = DLookup("商品コード", TBL_M_PROD, "商品ID = " & wrs.varRecordset.Fields("商品コード").VALUE)
            .Cells(product_table_1st_row_num + cnt_num, product_table_1st_column_num + 3).VALUE = DLookup("商品名", TBL_M_PROD, "商品ID = " & wrs.varRecordset.Fields("商品コード").VALUE)
            .Cells(product_table_1st_row_num + cnt_num, product_table_1st_column_num + 6).VALUE = wrs.varRecordset.Fields("数量").VALUE
            .Cells(product_table_1st_row_num + cnt_num, product_table_1st_column_num + 7).VALUE = DLookup("UNIT", TBL_M_BSC_UNIT, "ID = " & wrs.varRecordset.Fields("数量単位").VALUE)
            .Cells(product_table_1st_row_num + cnt_num, product_table_1st_column_num + 8).VALUE = DLookup("単価", TBL_M_PROD, "商品ID = " & wrs.varRecordset.Fields("商品コード").VALUE)
            .Cells(product_table_1st_row_num + cnt_num, product_table_1st_column_num + 12).VALUE = wrs.varRecordset.Fields("備考").VALUE
            If cnt_num >= 18 Then Exit Do
            cnt_num = cnt_num + 1
            wrs.varRecordset.MoveNext
        Loop
        
    End With

TruePoint:

    loadEstimationTable = True

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
'*機能      ：T_見積書項目データテーブルのレコード挿入
'*引数      ：データベース
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function insertEstmTblDtl(ByVal daoDB As dao.Database) As Boolean
    
    '定数
    Const FUNC_NAME As String = "insertEstmTblDtl"
    
    '変数
    Dim wrs As New clsWrappedRecordSet
    
    On Error GoTo ErrorHandler

    insertEstmTblDtl = False
    
    '既存の見積番号レコードを削除
    Call daoDB.execute("DELETE FROM " & TBL_T_ESTM_DTL & myVBVacant & "WHERE 見積番号 = " & myVBSglQte & Trim(myObjF05DtTrsfr.estiNum) & myVBSglQte & ";")
    
    '挿入
    Set wrs.varRecordset = daoDB.OpenRecordset(TBL_T_ESTM_DTL)
    With wrs.varRecordset
        .AddNew
        
        .Fields("見積番号").VALUE = Trim(myObjF05DtTrsfr.estiNum)
        .Fields("会社名").VALUE = myObjF05DtTrsfr.cmpNm
        .Fields("担当者").VALUE = myObjF05DtTrsfr.mmbr
        .Fields("納期").VALUE = Trim(myObjF05DtTrsfr.delivery)
        .Fields("支払条件").VALUE = Trim(myObjF05DtTrsfr.payCondition)
        .Fields("有効期限").VALUE = Trim(myObjF05DtTrsfr.validLimit)
        .Fields("その他備考").VALUE = Trim(myObjF05DtTrsfr.other)
        .Fields("消費税率").VALUE = myObjF05DtTrsfr.csmpTax
        .Fields("見積書作成日時").VALUE = CDate(CLng(Now))
        
        .Update
    End With

TruePoint:

    insertEstmTblDtl = True

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
'*機能      ：T_見積_表作成用データテーブルのレコード挿入
'*引数      ：データベース
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function insertEstmTbl(ByVal daoDB As dao.Database) As Boolean
    
    '定数
    Const FUNC_NAME As String = "insertEstmTbl"
    
    '変数
    Dim estiID As Long
    Dim arrWrs(0 To 1) As New clsWrappedRecordSet
    
    On Error GoTo ErrorHandler

    insertEstmTbl = False
    
    '既存の見積番号レコードを削除
    estiID = DLookup("見積ID", TBL_T_ESTM_DTL, "見積番号 = " & myVBSglQte & Trim(myObjF05DtTrsfr.estiNum) & myVBSglQte)
    Call daoDB.execute("DELETE FROM " & TBL_T_ESTM & myVBVacant & "WHERE 見積番号 = " & estiID & ";")
    
    '挿入
    Set arrWrs(0).varRecordset = daoDB.OpenRecordset(TBL_T_ESTM)
    Set arrWrs(1).varRecordset = daoDB.OpenRecordset("SELECT * FROM " & TBL_W_ESTM & " ORDER BY Number")
    With arrWrs(1).varRecordset
        .MoveFirst
        Do Until .EOF
            arrWrs(0).varRecordset.AddNew
            
            arrWrs(0).varRecordset.Fields("見積番号").VALUE = estiID
            arrWrs(0).varRecordset.Fields("商品コード").VALUE = .Fields("商品コード").VALUE
            arrWrs(0).varRecordset.Fields("数量").VALUE = .Fields("数量").VALUE
            arrWrs(0).varRecordset.Fields("数量単位").VALUE = .Fields("数量単位").VALUE
            arrWrs(0).varRecordset.Fields("備考").VALUE = .Fields("備考").VALUE
            
            arrWrs(0).varRecordset.Update
            .MoveNext
        Loop
    End With
    
    
TruePoint:

    insertEstmTbl = True

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

