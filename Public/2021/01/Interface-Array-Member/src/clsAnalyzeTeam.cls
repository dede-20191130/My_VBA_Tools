VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAnalyzeTeam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements clsAbsTeam

'**************************
'*チームクラス　解析チーム
'**************************

'定数欄

'変数欄
Private myArrayMenberName(1 To 6) As String



'******************************************************************************************
'*getter/setter欄
'******************************************************************************************
Private Property Let clsAbsTeam_arrayMenberName(ByVal idx As Long, ByVal name As String)
    myArrayMenberName(idx) = name
End Property

Private Property Get clsAbsTeam_arrayMenberName(ByVal idx As Long) As String
    clsAbsTeam_arrayMenberName = myArrayMenberName(idx)
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
    
    clsAbsTeam_getMemberName = idx & "番目のチームメンバーは" & myArrayMenberName(idx) & "です。"
    
    
ExitHandler:

    Exit Function
        
End Function

