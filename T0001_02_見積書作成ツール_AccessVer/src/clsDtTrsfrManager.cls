VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDtTrsfrManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Class")
Option Compare Database
Option Explicit

'**************************
'*データ転送クラス管理クラス
'**************************

'定数欄
Private Const SOURCE_NAME As String = "clsDtTrsfrManager"

'変数欄
Private dicSetObjDtTrsfr As Object


'******************************************************************************************
'*getter/setter欄
'******************************************************************************************




'******************************************************************************************
'*機能      ：Class_Initialize
'*引数      ：
'******************************************************************************************
Private Sub Class_Initialize()

    '定数
    Const FUNC_NAME As String = "Class_Initialize"
    
    '変数
    
    On Error GoTo ErrorHandler
    
    '初期化
    Set dicSetObjDtTrsfr = CreateObject(SCRIPTING_DICTIONARY)

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
'*機能      ：データ転送オブジェクトのキープ
'*引数      ：対象オブジェクト
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function setObjDtTrsfr(ByVal objDtTrsfr As Object) As Boolean
    
    '定数
    Const FUNC_NAME As String = "setObjDtTrsfr"
    
    '変数
    
    On Error GoTo ErrorHandler

    setObjDtTrsfr = False
    
    '内部にデータ転送オブジェクトを保持
    If Not dicSetObjDtTrsfr.Exists(TypeName(objDtTrsfr)) Then
        dicSetObjDtTrsfr.Add TypeName(objDtTrsfr), objDtTrsfr
    End If

TruePoint:

    setObjDtTrsfr = True

ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & err.Number & vbNewLine & _
           err.description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Function





'******************************************************************************************
'*機能      ：データ転送オブジェクトの取得
'*引数      ：目的オブジェクトの種類
'*戻り値    ：対象オブジェクト
'******************************************************************************************
Public Function getObjDtTrsfr(ByVal strDtTrsfr As String) As Object
    
    '定数
    Const FUNC_NAME As String = "getObjDtTrsfr"
    
    '変数
    
    Set getObjDtTrsfr = dicSetObjDtTrsfr.Item(strDtTrsfr)
    dicSetObjDtTrsfr.Remove (strDtTrsfr)


ExitHandler:

    Exit Function
    
End Function

