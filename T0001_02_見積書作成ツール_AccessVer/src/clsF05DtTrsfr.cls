VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsF05DtTrsfr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Class")
Option Compare Database
Option Explicit

'**************************
'*見積書項目設定Formデータ転送クラス
'**************************

'定数欄
Private Const SOURCE_NAME As String = "clsF05DtTrsfr"

'変数欄
Private myEstiNum As String
Private myCmpNm As Long
Private myMmbr As Long
Private myDelivery As String
Private myPayCondition As String
Private myValidLimit As String
Private myOther As String
Private myCsmpTax As Double



'******************************************************************************************
'*getter/setter欄
'******************************************************************************************


Public Property Let estiNum(ByVal RHS As String)
    myEstiNum = RHS
End Property

Public Property Get estiNum() As String
    estiNum = myEstiNum
End Property

Public Property Let cmpNm(ByVal RHS As Long)
    myCmpNm = RHS
End Property

Public Property Get cmpNm() As Long
    cmpNm = myCmpNm
End Property

Public Property Let mmbr(ByVal RHS As Long)
    myMmbr = RHS
End Property

Public Property Get mmbr() As Long
    mmbr = myMmbr
End Property

Public Property Let delivery(ByVal RHS As String)
    myDelivery = RHS
End Property

Public Property Get delivery() As String
    delivery = myDelivery
End Property

Public Property Let payCondition(ByVal RHS As String)
    myPayCondition = RHS
End Property

Public Property Get payCondition() As String
    payCondition = myPayCondition
End Property

Public Property Let validLimit(ByVal RHS As String)
    myValidLimit = RHS
End Property

Public Property Get validLimit() As String
    validLimit = myValidLimit
End Property

Public Property Let other(ByVal RHS As String)
    myOther = RHS
End Property

Public Property Get other() As String
    other = myOther
End Property

Public Property Let csmpTax(ByVal RHS As Double)
    myCsmpTax = RHS
End Property

Public Property Get csmpTax() As Double
    csmpTax = myCsmpTax
End Property
