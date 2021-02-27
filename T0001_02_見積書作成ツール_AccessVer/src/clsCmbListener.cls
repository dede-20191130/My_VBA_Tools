VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCmbListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Class")
Option Compare Database
Option Explicit

'**************************
'*コンボボックスイベント管理クラス
'**************************

'定数

'変数
Private WithEvents myComboBox As Access.ComboBox

'******************************************************************************************
'*getter/setter
'******************************************************************************************
Public Property Set ComboBox(ByRef cmb As Access.ComboBox)
    Set myComboBox = cmb
    myComboBox.OnKeyDown = "[Event Procedure]"
End Property

#If CBool(DEBUG_MODE) Then
Public Property Get ComboBox() As Access.ComboBox
    Set ComboBox = myComboBox
End Property
#End If



'******************************************************************************************
'*関数名    ：myComboBox_KeyDown
'*機能      ：キー入力を無効にする
'*引数(1)   ：
'******************************************************************************************
Private Sub myComboBox_KeyDown(KeyCode As Integer, Shift As Integer)
    
    '定数
    Const FUNC_NAME As String = "myComboBox_KeyDown"
    
    '変数
    
    On Error GoTo ErrorHandler
    
    
    'キー入力を無効にする(Enter/Tab/Esc以外)
    If KeyCode = vbKeyReturn Then GoTo ExitHandler
    If KeyCode = vbKeyTab Then GoTo ExitHandler
    If KeyCode = vbKeyEscape Then GoTo ExitHandler
    
    KeyCode = 0
    
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





