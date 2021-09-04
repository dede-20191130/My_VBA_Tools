Attribute VB_Name = "テストモジュール"
'@Folder("Database")
Option Compare Database
Option Explicit

Private Const SOURCE_NAME = "テストモジュール"



'******************************************************************************************
'*機能      ：テスト　コンボボックスの項目リストを変更関数
'******************************************************************************************
Public Sub テスト_changeCmbBoxItems()
    
    '定数
    Const FUNC_NAME As String = "テスト_changeCmbBoxItems"
    
    '変数
    Dim tForm As Form
    Dim fName As String
    Dim cmb As ComboBox
    
    On Error GoTo ErrorHandler

    '//フォームの動的作成
    Set tForm = CreateForm()
    fName = tForm.Name
    
    '//デザインビューで開く
    DoCmd.OpenForm fName, acDesign
    
    '//コンボボックスの動的作成
    Set cmb = CreateControl(fName, _
                            AcControlType.acComboBox)
    Dim mycmb As String
    mycmb = "mycmb"
    cmb.Name = mycmb
    cmb.RowSourceType = "Value List"
    
    '//デザインビューを閉じる
    DoCmd.Close acForm, fName, acSaveYes
    
    '//フォームビューで開く
    DoCmd.OpenForm fName, acNormal
    
    '//上記で作成したコンボボックスを再度参照
    Set cmb = Forms(fName).Controls(mycmb)
    
    '//■テスト01：食べ物のリスト設定
    '////関数呼び出し
    Call changeCmbBoxItems(1, cmb)
    '////アサーション
    Debug.Assert cmb.ListCount = 3
    Debug.Assert cmb.Column(0, 0) = "ピザ"
    Debug.Assert cmb.Column(0, 1) = "そば"
    Debug.Assert cmb.Column(0, 2) = "焼き肉"
    Debug.Print cmb.ListCount
        
    '//■テスト02：飲み物のリスト設定
    '////関数呼び出し
    Call changeCmbBoxItems(2, cmb)
    '////アサーション
    Debug.Assert cmb.ListCount = 3
    Debug.Assert cmb.Column(0, 0) = "コーラ"
    Debug.Assert cmb.Column(0, 1) = "緑茶"
    Debug.Assert cmb.Column(0, 2) = "水"
    Debug.Print cmb.ListCount
    
    '//フォームビューを閉じる
    DoCmd.Close , , acSaveNo
    
    '//動的生成したフォームを削除
    DoCmd.DeleteObject acForm, fName
    
ExitHandler:
    
    '//テスト完了
    Debug.Print Now & ":Finish " & FUNC_NAME
    
    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "クラス名：" & SOURCE_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical
        
    GoTo ExitHandler
        
End Sub



'******************************************************************************************
'*機能      ：テスト　テキストボックスの使用可能状態の変更関数
'******************************************************************************************
Public Sub テスト_changeTextBoxesEnabled()
    
    '定数
    Const FUNC_NAME As String = "テスト_changeTextBoxesEnabled"
    
    '変数
    Dim tForm As Form
    Dim fName As String
    Dim textboxes(0 To 3) As textbox
    Dim i As Long
    
    On Error GoTo ErrorHandler

    '//フォームの動的作成
    Set tForm = CreateForm()
    fName = tForm.Name
    
    '//デザインビューで開く
    DoCmd.OpenForm fName, acDesign
    
    '//テキストボックス配列の動的作成
    For i = 0 To 3
        Set textboxes(i) = CreateControl(fName, _
                            AcControlType.acTextBox)
                            
        textboxes(i).Name = "mytext_" & i
        
        '//一部のみunder18、それ以外はover18のタグを付与
        If i < 2 Then
            textboxes(i).Tag = "under18"
        Else
            textboxes(i).Tag = "over18"
        End If
    Next i
    
    '//デザインビューを閉じる
    DoCmd.Close acForm, fName, acSaveYes
    
    '//フォームビューで開く
    DoCmd.OpenForm fName, acNormal
    
    '//上記で作成したテキストボックス配列を再度参照
    For i = 0 To 3
        Set textboxes(i) = Forms(fName).Controls("mytext_" & i)
    Next i
    
    '//■テスト01：18歳未満専用のテキストボックスの有効化
    '////関数呼び出し
    Call changeTextBoxesEnabled(1, textboxes)
    '////アサーション
    Debug.Assert textboxes(0).Tag = "under18"
    Debug.Assert textboxes(0).Enabled = True
    Debug.Assert textboxes(1).Tag = "under18"
    Debug.Assert textboxes(1).Enabled = True
    Debug.Assert textboxes(2).Tag <> "under18"
    Debug.Assert textboxes(2).Enabled = False
    Debug.Assert textboxes(3).Tag <> "under18"
    Debug.Assert textboxes(3).Enabled = False
    
    '//■テスト02：18歳以上専用のテキストボックスの有効化
    '////関数呼び出し
    Call changeTextBoxesEnabled(2, textboxes)
    '////アサーション
    Debug.Assert textboxes(0).Tag = "under18"
    Debug.Assert textboxes(0).Enabled = False
    Debug.Assert textboxes(1).Tag = "under18"
    Debug.Assert textboxes(1).Enabled = False
    Debug.Assert textboxes(2).Tag <> "under18"
    Debug.Assert textboxes(2).Enabled = True
    Debug.Assert textboxes(3).Tag <> "under18"
    Debug.Assert textboxes(3).Enabled = True
        
    '//フォームビューを閉じる
    DoCmd.Close , , acSaveNo
    
    '//動的生成したフォームを削除
    DoCmd.DeleteObject acForm, fName
    
ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "クラス名：" & SOURCE_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical
        
    GoTo ExitHandler
        
End Sub

