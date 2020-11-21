Attribute VB_Name = "M_Debug"
'@Folder("Database")
Option Compare Database
Option Explicit


'**************************
'*動作確認用
'**************************



#If DEBUG_MODE Then


Sub s20201118_2314()
    On Error Resume Next
    Dim obj As clsCreateNewExcel: Set obj = New clsCreateNewExcel
    Dim ws
    
    obj.WorkSheets(1).cells(1, 1).Value = 123
    
    obj.addNewSheet
    obj.addNewSheet
    
    '    Set ws = obj.WorkSheets(4)
    
    obj.WorkSheets(3).Range("B2").interior.Color = RGB(255, 100, 30)
    
    Set ws = obj.addNewSheet
    
    Debug.Assert ws.Name = obj.WorkSheets(4).Name
    
    Debug.Print Err.Description
    
End Sub


Sub s20201119_2115()
    Dim arr() As String
    
    Set M_Export.db = CurrentDb
    Debug.Print getTableHeader("M_商品データ", arr)
    
End Sub


Sub s20201119_2227()
    Dim arr() As Variant
    
    Set M_Export.db = CurrentDb
    
    Debug.Print getTableDataBySQL("select * from M_商品データ where (商品id mod 3) = 2;", arr)
    
End Sub


Sub s20201120_0235()
    Dim jsonStr As String
    Dim obj As Dictionary
    Dim arr() As Variant
    Dim i As Long
    Dim j As Long
    
    jsonStr = RESTAPI20201120_0243(WEBAPI_URL)
    
    Set obj = JsonConverter.ParseJson(jsonStr)
    
    ReDim arr(0 To obj.Count - 1, 0 To 2)
    For i = LBound(obj.Keys) To UBound(obj.Keys)
        arr(i, 0) = obj.Keys(i)
        arr(i, 1) = obj.Item(arr(i, 0)).Item("name")
        arr(i, 2) = obj.Item(arr(i, 0)).Item("price")
    Next
    
End Sub


Function RESTAPI20201120_0243(URL As String)
    Dim objXMLHttp As Object
    Set objXMLHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    objXMLHttp.Open "GET", URL, False
    objXMLHttp.Send

    RESTAPI20201120_0243 = objXMLHttp.responseText
End Function


#End If
