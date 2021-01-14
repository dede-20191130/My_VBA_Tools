VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsNewTeam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements clsAbsTeam

'**************************
'*チームクラス　新設のチーム
'**************************

'定数欄

'変数欄


'******************************************************************************************
'*getter/setter欄
'******************************************************************************************
Private Property Let clsAbsTeam_arrayMenberName(ByVal idx As Long, ByVal name As String)
    '何もしない
End Property

Private Property Get clsAbsTeam_arrayMenberName(ByVal idx As Long) As String
    clsAbsTeam_arrayMenberName = "名前無し"
End Property

'******************************************************************************************
'*関数名    ：clsAbsTeam_getMemberName
'*機能      ：チームメンバーの名前取得
'*引数      ：対象者のインデックス番号
'*戻り値    ：チームメンバーの名前
'******************************************************************************************
Private Function clsAbsTeam_getMemberName(ByVal idx As Long) As String
    
    '定数
    
    '変数
    
    clsAbsTeam_getMemberName = "※新設のチームにはメンバーが存在しません。"
    
    
ExitHandler:

    Exit Function
        
End Function


