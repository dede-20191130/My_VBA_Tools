VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTxtListenerNumberBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Class")
Option Compare Database
Option Explicit

'**************************
'*数値入力テキストボックスイベント管理クラス
'**************************

'定数

'変数
Private WithEvents myTextBox As Access.TextBox
Attribute myTextBox.VB_VarHelpID = -1

'******************************************************************************************
'*getter/setter
'******************************************************************************************
Public Property Set TextBox(ByRef cmb As Access.TextBox)
    Set myTextBox = cmb
    myTextBox.OnKeyPress = "[Event Procedure]"
End Property

#If CBool(DEBUG_MODE) Then
Public Property Get TextBox() As Access.TextBox
    Set TextBox = myTextBox
End Property
#End If


'******************************************************************************************
'*関数名    ：myTextBox_KeyPress
'*機能      ：数値のみ入力可能
'*引数(1)   ：アスキーコード
'******************************************************************************************
Private Sub myTextBox_KeyPress(KeyAscii As Integer)
    
    '定数
    Const FUNC_NAME As String = "myTextBox_KeyPress"
    
    '変数
    
    On Error GoTo ErrorHandler
    
    'KeyAsciiが特定の値のみ有効
    'i. 0と9の間
    'ii.Backspaceおよびピリオド
    If 47 < KeyAscii And KeyAscii < 58 Then GoTo ExitHandler
    If KeyAscii = 8 Then GoTo ExitHandler
    If KeyAscii = 46 Then GoTo ExitHandler
    
    KeyAscii = 0
    
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

