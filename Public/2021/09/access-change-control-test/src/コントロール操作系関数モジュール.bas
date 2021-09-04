Attribute VB_Name = "コントロール操作系関数モジュール"
'@Folder("Database")
Option Compare Database
Option Explicit

Private Const SOURCE_NAME = "コントロール操作系関数モジュール"

'******************************************************************************************
'*機能      ：コンボボックスの項目リストを変更
'*引数      ：
'******************************************************************************************
Public Sub changeCmbBoxItems(ByVal selectedNumber As Long, ByVal cmbBox As ComboBox)
    
    '定数
    Const FUNC_NAME As String = "changeCmbBoxItems"
    
    '変数
    
    On Error GoTo ErrorHandler
    
    '//項目のクリア
    cmbBox.RowSource = ""
    
    Select Case selectedNumber
    '//食べ物
    Case 1
        cmbBox.AddItem "ピザ"
        cmbBox.AddItem "そば"
        cmbBox.AddItem "焼き肉"
    '//飲み物
    Case 2
        cmbBox.AddItem "コーラ"
        cmbBox.AddItem "緑茶"
        cmbBox.AddItem "水"
    End Select

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





'******************************************************************************************
'*機能      ：テキストボックスの使用可能状態を変更
'*引数      ：
'******************************************************************************************
Public Sub changeTextBoxesEnabled(ByVal selectedNumber As Long, ByRef textboxes() As textbox)
    
    '定数
    Const FUNC_NAME As String = "changeTextBoxesEnabled"
    
    '変数
    Dim canUnder18Enable As Boolean '//18歳未満のためのテキストボックスが有効かどうか
    Dim textbox As Variant
    
    On Error GoTo ErrorHandler
    
    '//18歳未満を選択時はTrue、それ以外の場合はFalse
    canUnder18Enable = selectedNumber = 1
    
    '//タグがunder18かover18かによって
    '//使用可能状態を切り替える
    For Each textbox In textboxes
        If InStr(textbox.Tag, "under18") <> 0 Then
            textbox.Enabled = canUnder18Enable
        Else
            textbox.Enabled = Not canUnder18Enable
        End If
    Next textbox

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


