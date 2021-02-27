Attribute VB_Name = "mdlManageForm"
'@Folder("Database")
Option Compare Database
Option Explicit

'**************************
'*フォーム管理モジュール
'**************************

'定数欄
Private Const SOURCE_NAME As String = "mdlManageForm"

'変数欄



'******************************************************************************************
'*getter/setter欄
'******************************************************************************************




'******************************************************************************************
'*機能      ：非可視化フォームor未ロードフォームを開く
'*引数      ：対象フォーム名前
'*引数      ：データ受け渡し引数
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function showFormInvisibleOrUnloaded(ByVal pFormName As String, Optional pOpenArgs As String = "") As Boolean
    
    '定数
    Const FUNC_NAME As String = "YshowFormInvisibleOrUnloadedYY"
    
    '変数
    
    On Error GoTo ErrorHandler

    showFormInvisibleOrUnloaded = False
    
    '現在読み込まれているフォームならば再可視化
    If Application.CurrentProject.AllForms(pFormName).IsLoaded Then
        Forms(pFormName).Visible = True
    'フォームを開く
    Else
        DoCmd.OpenForm pFormName, , , , , , _
            pOpenArgs
    End If

TruePoint:

    showFormInvisibleOrUnloaded = True

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
'*機能      ：ロード済みのフォームを閉じる
'*引数      ：可変長　対象のフォーム
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function closeFormIfLoaded(ParamArray pArrFormName() As Variant) As Boolean
    
    '定数
    Const FUNC_NAME As String = "closeFormIfLoaded"
    
    '変数
    
    On Error GoTo ErrorHandler

    closeFormIfLoaded = False
    
    Dim c As Variant
    For Each c In pArrFormName
        If Application.CurrentProject.AllForms(CStr(c)).IsLoaded Then DoCmd.Close acForm, CStr(c), acSaveNo
    Next c
    

TruePoint:

    closeFormIfLoaded = True

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

