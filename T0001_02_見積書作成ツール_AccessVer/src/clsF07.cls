VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsF07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Class")
Option Compare Database
Option Explicit

'**************************
'*見積書項目レコードの情報選択・取得Form処理クラス
'**************************

'定数欄
Private Const SOURCE_NAME As String = "clsF07"

'変数欄



'******************************************************************************************
'*getter/setter欄
'******************************************************************************************




'******************************************************************************************
'*機能      ：btnOKClick処理
'*引数      ：
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function btnOKClick(ByVal pForm As Object) As Boolean
    
    '定数
    Const FUNC_NAME As String = "btnOKClick"
    
    '変数
    Dim boolErr As Boolean
    Dim collNoExist As New Collection
    
    On Error GoTo ErrorHandler

    btnOKClick = False
    
    '空欄チェック
    If Not checkWhetherControlsVacant(boolErr, pForm.lst_selectField.Column(1)) Then GoTo ExitHandler
    If boolErr Then Call MsgBox("リストのレコードを選択してください。", vbExclamation, MESSAGE_TITLE_WARNING): GoTo TruePoint

    '呼び出し元に選択値を反映
    If pForm.chk_company.VALUE Then
        Dim orgID As Variant: orgID = DLookup("[会社ID]", TBL_M_ORG, "[会社名] = '" & Nz(pForm.lst_selectField.Column(2), "") & "'")
        If Not IsNull(orgID) Then
            Form_F05_見積書項目設定.lst_cmpNm.RowSource = "SELECT 会社ID,会社名 FROM " & _
                TBL_M_ORG & " " & _
                "WHERE 会社ID = " & _
                orgID & ";"
        Else
            collNoExist.Add Form_F05_見積書項目設定.lst_cmpNm.Tag
        End If
    End If
    If pForm.chk_member.VALUE Then
        Dim memID As Variant: memID = DLookup("[人名ID]", TBL_M_MEMBER, "[名前] = '" & Nz(pForm.lst_selectField.Column(3), "") & "'")
        If Not IsNull(memID) Then
            Form_F05_見積書項目設定.lst_mmbr.RowSource = "SELECT 人名ID,名前 FROM " & _
                TBL_M_MEMBER & " " & _
                "WHERE 人名ID = " & _
                memID & ";"
        Else
            collNoExist.Add Form_F05_見積書項目設定.lst_mmbr.Tag
        End If
    End If
    If pForm.chk_delibery.VALUE Then Form_F05_見積書項目設定.txt_delivery.VALUE = Nz(pForm.lst_selectField.Column(4), "")
    If pForm.chk_payCondition.VALUE Then Form_F05_見積書項目設定.txt_payCondition.VALUE = Nz(pForm.lst_selectField.Column(5), "")
    If pForm.chk_validLimit.VALUE Then Form_F05_見積書項目設定.txt_validLimit.VALUE = Nz(pForm.lst_selectField.Column(6), "")
    If pForm.chk_other.VALUE Then Form_F05_見積書項目設定.txt_other.VALUE = DLookup("[その他備考]", TBL_T_ESTM_DTL, "[見積ID] = " & Nz(pForm.lst_selectField.Column(0), 0))
    If pForm.chk_csmpTax.VALUE Then
        Dim tax As Variant: tax = DLookup("[ID]", TBL_M_BSC_TAX, "[value] = " & Nz(pForm.lst_selectField.Column(8), 0))
        If Not IsNull(tax) Then
            Form_F05_見積書項目設定.cmb_csmpTax.VALUE = tax
        Else
            collNoExist.Add Form_F05_見積書項目設定.cmb_csmpTax.Tag
        End If
    End If
    
    'マスタに存在しない項目の反映不可の警告
    If collNoExist.Count > 0 Then
        MsgBox "下記項目はマスタから削除済みのため反映されませんでした。" & vbNewLine & _
                    vbNewLine & _
                    Join(CollectionToArray(collNoExist), vbNewLine), _
                vbExclamation, _
                MESSAGE_TITLE_WARNING
    End If
    
    '画面を閉じる
    DoCmd.Close acForm, pForm.Name, acSaveNo
    
TruePoint:

    btnOKClick = True

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

