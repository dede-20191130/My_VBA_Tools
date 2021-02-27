VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAbsSettingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Class")
Option Compare Database
Option Explicit

'**************************
'*セッティングForm処理クラスインタフェース
'**************************

'定数欄
Private Const SOURCE_NAME As String = "clsAbsSettingForm"

'変数欄



'******************************************************************************************
'*getter/setter欄
'******************************************************************************************
Public Property Get objSettingFormCommon() As clsSettingFormCommon

End Property

Public Property Get typeSettingForm() As eTypeSettingForm
    
End Property


Public Function getInfo(ByRef pDicInfo As Object) As Boolean

End Function
