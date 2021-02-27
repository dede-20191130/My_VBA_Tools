VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormSettingSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Class")
Option Compare Database
Option Explicit

'**************************
'*設定フォームサブフォームイベント管理クラス
'**************************

'定数

'変数
Private WithEvents myForm As Access.Form
Attribute myForm.VB_VarHelpID = -1


'******************************************************************************************
'*getter/setter
'******************************************************************************************
Public Property Set Form(ByRef frm As Access.Form)
    Set myForm = frm
    myForm.OnCurrent = "[Event Procedure]"
End Property

#If CBool(DEBUG_MODE) Then
Public Property Get Form() As Access.Form
    Set Form = myForm
End Property
#End If



'******************************************************************************************
'*関数名    ：myForm_Current
'*機能      ：
'*引数      ：
'******************************************************************************************
Private Sub myForm_Current()
    
    '定数
    Const FUNC_NAME As String = "myForm_Current"
    
    '変数
    
    On Error GoTo ErrorHandler

    'メインフォームの選択コンボボックスの値が
    'フォームの現在のIDと同じであれば終了
    If myForm.Parent.cmb_selectedRcd.VALUE = myForm.txt_ID.VALUE Then GoTo ExitHandler
    
    'メインフォームの選択コンボボックスの値をIDと同じにする
    myForm.Parent.cmb_selectedRcd.VALUE = myForm.txt_ID.VALUE
    
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


