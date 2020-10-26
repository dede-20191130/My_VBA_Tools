Attribute VB_Name = "M_UI"
'@Folder("Database")
Option Compare Database
Option Explicit

'**************************
'*ユーザインタフェースModule
'**************************

'定数
Public Const msoFileDialogFilePicker As Long = 3


'変数


'******************************************************************************************
'*関数名    ：getFilePathFromDialog
'*機能      ：ダイアログで選択されたファイルパスを取得
'*引数(1)   ：タイトル
'*引数(2)   ：フィルタに指定するkey/valueの辞書
'*戻り値    ：ファイルパス
'******************************************************************************************
Public Function getFilePathFromDialog( _
       Optional ByVal pTitle As String = "選択ダイアログ", _
       Optional ByVal dicFilter As Object = Nothing _
       ) As String
    
    '定数
    Const FUNC_NAME As String = "getFilePathFromDialog"
    
    '変数
    Dim cntVal As Variant
    Dim filePath As String
    
    On Error GoTo ErrorHandler

    getFilePathFromDialog = ""
    
    'ダイアログ設定
    With Application.FileDialog(msoFileDialogFilePicker)
    
        .Title = pTitle
        
        .Filters.Clear
        If Not dicFilter Is Nothing Then
            For Each cntVal In dicFilter.Keys
                .Filters.Add cntVal, dicFilter.Item(cntVal)
            Next cntVal
            .FilterIndex = 1
        End If
        
        '複数ファイル選択の禁止
        .AllowMultiSelect = False
                
        'キャンセル時は終了
        If .Show <> -1 Then GoTo ExitHandler
        
        '選択されたファイルパス
        filePath = .SelectedItems(1)
                
    End With

    getFilePathFromDialog = filePath
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.description, vbCritical, Tool_Name
        
    GoTo ExitHandler
        
End Function



