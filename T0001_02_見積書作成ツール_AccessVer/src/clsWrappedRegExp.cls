VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWrappedRegExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Class")
Option Compare Database
Option Explicit

'**************************
'*正規表現利用クラス
'**************************

'定数欄
Private Const SOURCE_NAME As String = "clsWrappedRegExp"

'変数欄
Private objRegex As RegExp


'******************************************************************************************
'*getter/setter欄
'******************************************************************************************
Public Property Let regGlobal(ByVal RHS As Boolean)
    objRegex.Global = RHS
End Property

Public Property Let regIgnoreCase(ByVal RHS As Boolean)
    objRegex.IgnoreCase = RHS
End Property


'******************************************************************************************
'*機能      ：コンストラクタ
'*引数      ：
'******************************************************************************************
Private Sub Class_Initialize()

    '定数
    Const FUNC_NAME As String = "Class_Initialize"
    
    '変数
    
    On Error GoTo ErrorHandler
    
    Set objRegex = New RegExp
    With objRegex
        .Global = True               '// 検索範囲（True：文字列の最後まで検索、False：最初の一致まで検索）
        .IgnoreCase = True           '// 大文字小文字の区別（True：区別しない、False：区別する）
    End With

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & err.Number & vbNewLine & _
           err.description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Sub


'******************************************************************************************
'*機能      ：デストラクタ
'*引数      ：
'******************************************************************************************
Private Sub Class_Terminate()

    '定数
    Const FUNC_NAME As String = "Class_Terminate"
    
    '変数
    
    On Error GoTo ErrorHandler

    Set objRegex = Nothing

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & err.Number & vbNewLine & _
           err.description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Sub


 
'******************************************************************************************
'*機能      ：パターン一致判定
'*              指定された文字列を正規表現で検索し、パターンに一致する文字列が
'             検索されたかどうかを示すブール (Boolean) 値を返却。
'*引数      ：target[String] ：検索対象文字列
'*引数      ：pattern[String]：検索する文字列（正規表現パターン）
'*戻り値    ：True > 一致、False > 不一致
'******************************************************************************************
Public Function test(ByVal target As String, ByVal pattern As String) As Boolean
    
    '定数
    Const FUNC_NAME As String = "test"
    
    '変数
    
    objRegex.pattern = pattern
    test = objRegex.test(target)

ExitHandler:

    Exit Function
    
End Function

 
'******************************************************************************************
'*機能      ：正規表現による検索で見つかったテキストを置換
'*引数      ：target[String]：検索および置換の対象となるテキスト文字列
'*引数      ：before[String]：検索する文字列（正規表現パターン）
'*引数      ：after[String] ：置換するテキスト文字列
'*戻り値    ：置換後文字列
'******************************************************************************************
Public Function replace(ByVal target As String, ByVal before As String, ByVal after As String) As String
    
    '定数
    Const FUNC_NAME As String = "Replace"
    
    '変数
    
    objRegex.pattern = before
    replace = objRegex.replace(target, after)


ExitHandler:

    Exit Function
    
End Function


'******************************************************************************************
'*機能      ：指定された文字列を正規表現で検索
'*引数      ：target[String] ：検索対象文字列
'*引数      ：pattern[String]：検索する文字列（正規表現パターン）
'*戻り値    ：Matchesコレクション
'******************************************************************************************
Public Function execute(ByVal target As String, ByVal pattern As String) As MatchCollection
    
    
    '定数
    Const FUNC_NAME As String = "execute"
    
    '変数
    
    objRegex.pattern = pattern
    Set execute = objRegex.execute(target)


ExitHandler:

    Exit Function
    
End Function
