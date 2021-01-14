Attribute VB_Name = "M_Calling"
Option Explicit

'**************************
'*呼び出し元モジュール
'**************************


'******************************************************************************************
'*関数名    ：output3rdMemberName
'*機能      ：動作テスト
'*引数      ：
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Sub testFunc()
    
    '定数
    Const FUNC_NAME As String = "testFunc"
    
    '変数
    Dim team As clsAbsTeam
    Dim coll As New Collection
    
    On Error GoTo ErrorHandler

    '解析チームの名前を設定する
    Set team = New clsAnalyzeTeam
    team.arrayMenberName(1) = "佐藤"
    team.arrayMenberName(3) = "Mike"
    team.arrayMenberName(5) = "梦蝶"
    
    '処理するチームを追加
    coll.Add team
    coll.Add New clsNewTeam
    
    '名前を出力
    If Not outputSelectedMemberName(coll, 3) Then GoTo ExitHandler

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Interface-Array-Member"
        
    GoTo ExitHandler
        
End Sub



'******************************************************************************************
'*関数名    ：outputSelectedMemberName
'*機能      ：引数のコレクションの各チームの「idx」番目のメンバー名前を出力する
'*引数      ：
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Private Function outputSelectedMemberName(ByVal collTeam As Collection, ByVal idx As Long) As Boolean
    
    '定数
    Const FUNC_NAME As String = "outputSelectedMemberName"
    
    '変数
    Dim cntTeam As clsAbsTeam
    
    On Error GoTo ErrorHandler

    outputSelectedMemberName = False
    
    For Each cntTeam In collTeam
        Debug.Print cntTeam.getMemberName(idx)
    Next cntTeam

TruePoint:

    outputSelectedMemberName = True

ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Interface-Array-Member"
        
    GoTo ExitHandler
        
End Function

