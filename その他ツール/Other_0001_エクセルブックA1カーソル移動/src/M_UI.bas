Attribute VB_Name = "M_UI"
Option Explicit

'**************************
'*ユーザインタフェースModule
'**************************

'定数


'変数


'******************************************************************************************
'*関数名    ：getFilePathFromDialog
'*機能      ：ダイアログで選択されたフォルダパスを取得
'*引数(1)   ：タイトル
'*戻り値    ：フォルダパス
'******************************************************************************************
Public Function getFolderPathFromDialog( _
       Optional ByVal pTitle As String = "選択ダイアログ" _
       ) As String
    
    '定数
    Const FUNC_NAME As String = "getFilePathFromDialog"
    
    '変数
    Dim filePath As String
    
    On Error GoTo ErrorHandler

    getFolderPathFromDialog = ""
    
    'ダイアログ設定
    With Application.FileDialog(msoFileDialogFolderPicker)
    
        .Title = pTitle
        
        'キャンセル時は終了
        If .Show <> -1 Then GoTo ExitHandler
        
        '選択されたフォルダパス
        filePath = .SelectedItems(1)
                
    End With

    getFolderPathFromDialog = filePath
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Function


