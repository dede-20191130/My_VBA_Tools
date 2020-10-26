Attribute VB_Name = "M_Common"
'@Folder("Database")
Option Compare Database
Option Explicit

'**************************
'*共通関数Module
'**************************


'定数


'変数



#If VBA7 And Win64 Then
    '******************************************************************************************
    '*関数名    ：MessageBoxTimeoutA
    '*機能      ：自動で消去されるメッセージを表示
    '******************************************************************************************
    Public Declare PtrSafe Function MessageBoxTimeoutA Lib "User32" ( _
        ByVal Hwnd As Long, _
        ByVal lpText As String, _
        ByVal lpCaption As String, _
        ByVal uType As VbMsgBoxStyle, _
        ByVal wLanguageID As Long, _
        ByVal dwMilliseconds As Long) As Long
     
#Else
 
    Public Declare Function MessageBoxTimeoutA Lib "User32"( _
        ByVal Hwnd As Long, _
        ByVal lpText As String, _
        ByVal lpCaption As String, _
        ByVal uType As VbMsgBoxStyle, _
        ByVal wLanguageID As Long, _
        ByVal dwMilliseconds As Long) As Long
     
#End If

