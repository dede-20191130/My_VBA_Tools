Attribute VB_Name = "M_Export"
'@Folder("Database")
Option Compare Database
Option Explicit

'**************************
'*Dara Output Module
'**************************


'Vars
Public db As DAO.Database



'******************************************************************************************
'*Function :get header data of target table
'*Arg      :table name
'*Arg      :array for gotten data
'*Return   :True > normal termination; False > abnormal termination
'******************************************************************************************
Public Function getTableHeader(ByVal tblName As String, ByRef pArrHeader() As String) As Boolean
    
    'Consts
    Const FUNC_NAME As String = "getTableHeader"
    
    'Vars
    Dim i As Long
    
    On Error GoTo ErrorHandler

    getTableHeader = False
    Erase pArrHeader
    
    With db.TableDefs(tblName)
        ReDim pArrHeader(0 To .Fields.Count - 1)
        For i = 0 To .Fields.Count - 1
            pArrHeader(i) = .Fields(i).Name
        Next
    End With
    
    getTableHeader = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Function


'******************************************************************************************
'*Function :get recordset data as a 2-dimentional array
'*Arg      :sql string for target recordset
'*Arg      :array for gotten data
'*Return   :True > normal termination; False > abnormal termination
'******************************************************************************************
Public Function getTableDataBySQL(ByVal sql As String, ByRef arrData() As Variant) As Boolean
    
    'Consts
    Const FUNC_NAME As String = "getTableDataBySQL"
    
    'Vars
    Dim rs As DAO.Recordset
    Dim i As Long
    Dim j As Long
    
    On Error GoTo ErrorHandler

    getTableDataBySQL = False
    Erase arrData
    
    Set rs = db.OpenRecordset(sql)
    With rs
        If .EOF Then GoTo TruePoint
        .MoveLast
        ReDim arrData(0 To .RecordCount - 1, 0 To .Fields.Count - 1)
        .MoveFirst
        
        i = 0
        Do Until .EOF
            For j = 0 To .Fields.Count - 1
                arrData(i, j) = .Fields(j).Value
            Next j
            i = i + 1
            .MoveNext
        Loop
    End With

TruePoint:

    getTableDataBySQL = True
    
ExitHandler:
    
    If Not rs Is Nothing Then rs.Clone: Set rs = Nothing
    
    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Function




'******************************************************************************************
'*Function :post data to sheet
'*Arg      :target sheet
'*Arg      :assigned sheet name
'*Arg      :aheader data array
'*Arg      :data array
'*Return   :True > normal termination; False > abnormal termination
'******************************************************************************************
Public Function postDataToSheet( _
    ByVal tgtSheet As Object, _
    ByVal sheetName As String, _
    ByVal pArrHeader As Variant, _
    ByVal pArrData As Variant _
) As Boolean
    
    'Consts
    Const FUNC_NAME As String = "postDataToSheet"
    
    'Vars
    
    On Error GoTo ErrorHandler

    postDataToSheet = False
    
    With tgtSheet
        .Name = sheetName
        .Range(.cells(1, 1), .cells(1, UBound(pArrHeader) - LBound(pArrHeader) + 1)).Value = pArrHeader
        .Range(.cells(2, 1), .cells(UBound(pArrData, 1) - LBound(pArrData, 1) + 2, UBound(pArrData, 2) - LBound(pArrData, 2) + 1)).Value = pArrData
        'lines
        .Range(.cells(1, 1), .cells(UBound(pArrData, 1) - LBound(pArrData, 1) + 2, UBound(pArrData, 2) - LBound(pArrData, 2) + 1)).Borders.LineStyle = xlContinuous
        'column widths adjustment
        .Range(.Columns(1), .Columns(UBound(pArrHeader) - LBound(pArrHeader) + 1)).AutoFit
    End With

    postDataToSheet = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Function



'******************************************************************************************
'*Function :get Json string from specified URL
'*Arg      :URL
'*Return   :Json string
'******************************************************************************************
Public Function getJsonFromAPI(URL As String) As String

    'Consts
    Const FUNC_NAME As String = "getJsonFromAPI"
    
    'Vars
    Dim objXMLHttp As Object
    
    On Error GoTo ErrorHandler

    getJsonFromAPI = ""
    
    Set objXMLHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        objXMLHttp.Open "GET", URL, False
        objXMLHttp.Send


    getJsonFromAPI = objXMLHttp.responseText
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Function

