Attribute VB_Name = "M_ImpExp"
'@Folder("Database")
Option Compare Database
Option Explicit

'**************************
'*Import/Export Module
'**************************

'Const


'Variable


'******************************************************************************************
'*Function    :exportTableDefTablesMain
'*Function    :main function for export
'*Arg(1)      :target access file path
'*Return      :True > normal termination; False > abnormal termination


'******************************************************************************************
Public Function exportTableDefTablesMain(ByVal dbFilePath As String) As Boolean
    
    'Const
    Const FUNC_NAME As String = "exportTableDefTablesMain"
    
    'Variable
    Dim selectedDB As DAO.Database
    Dim xlApp As Object
    Dim wb As Object
    Dim tdf As DAO.TableDef
    Dim defArr As Variant
    Dim fstWs As Object
    Dim ws As Object
    
    On Error GoTo ErrorHandler
    
    exportTableDefTablesMain = False
    
    'connect the database in the access file
    Set selectedDB = getAccessDB(dbFilePath)
    If selectedDB Is Nothing Then GoTo ExitHandler
    
    'create new excel app instance and excel-book instance
    Set xlApp = CreateObject("Excel.Application")
    With xlApp
        .Visible = False
        .DisplayAlerts = False
        .ScreenUpdating = False
    End With
    Set wb = xlApp.Workbooks.Add
    
    Set fstWs = wb.Worksheets(1)
    
    'create a access table definition information table in separate sheets
    For Each tdf In selectedDB.TableDefs
        Do
            'do continue if tdf is one of the unnecessary tables such as system table.
            If Left(tdf.Name, 4) = "Msys" Or Left(tdf.Name, 4) = "Usys" Or Left(tdf.Name, 1) = "~" Then Exit Do
            
            'get definition information of the table
            defArr = getTableDefArray(tdf, selectedDB)
            If IsNull(defArr) Then GoTo ExitHandler
            
            'create a new sheet
            Set ws = wb.Worksheets.Add
            If Not setWSName(ws, tdf.Name) Then Call Err.Raise(1000, "Sheet Name Specification Error", "An error has occurred on sheet name specification.")
            
            'write a definition information to Range and auto-adjust the sheet column widths.
            With ws.Range(ws.cells(1, 1), ws.cells(UBound(defArr) - LBound(defArr) + 1, UBound(defArr, 2) - LBound(defArr, 2) + 1))
                .Value = defArr
                .EntireColumn.AutoFit
            End With
            
        Loop While False
    Next tdf
    
    'remove default sheet
    If wb.Worksheets.Count > 1 Then fstWs.Delete
    
    'save
    wb.saveas Left( _
              dbFilePath, _
              InStrRev(dbFilePath, ".") - 1 _
              ) & _
                "_Table_Info_List.xlsx"
    
    'Complete
    MessageBoxTimeoutA 0&, "Completed", "INFO", vbOKOnly, 0&, 3000
    
    exportTableDefTablesMain = True
    
ExitHandler:
    
    'Close
    If Not wb Is Nothing Then wb.Close: Set wb = Nothing
    If Not xlApp Is Nothing Then xlApp.Quit: Set xlApp = Nothing
    If Not selectedDB Is Nothing Then selectedDB.Close: Set selectedDB = Nothing
    
    Set tdf = Nothing
    Set ws = Nothing
    Set fstWs = Nothing
    
    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.description, vbCritical, Tool_Name
        
    GoTo ExitHandler
        
End Function


'******************************************************************************************
'*Function      : get definition information of the table
'*                Items:
'*                      Field Name
'*                      Data Type
'*                      Size
'*                      Required or not
'*                      Primary key or not
'*                      Foreign key or not
'*                      Description
'*Arg(1)        : TableDef instance
'*Arg(2)        : DAO Database instance
'*return        : definition information array

'******************************************************************************************
Public Function getTableDefArray( _
       ByVal pTdf As DAO.TableDef, _
       ByVal pSelectedDB As DAO.Database _
       ) As Variant
    
    'Const
    Const FUNC_NAME As String = "getTableDefArray"
    
    'Variable
    Dim defArr() As Variant
    Dim fld As DAO.Field
    Dim i As Long
    Dim dicPKs As Object
    Dim dicFKs As Object
    Dim description As String
    
    On Error GoTo ErrorHandler

    getTableDefArray = Null
    
    'do redimension the array to be (Field Count + 1) rows and 7 columns
    ReDim defArr(0 To pTdf.Fields.Count, 0 To 6)
    
    'setting the header part
    defArr(0, 0) = "Field Name"
    defArr(0, 1) = "Data Type"
    defArr(0, 2) = "Size"
    defArr(0, 3) = "Required or not"
    defArr(0, 4) = "Primary key or not"
    defArr(0, 5) = "Foreign key or not"
    defArr(0, 6) = "Description"
    
    'get a dictionary containing all primary key field names
    Set dicPKs = getPKs(pTdf)
    If dicPKs Is Nothing Then GoTo ExitHandler
    
    'get a dictionary containing all foreign key field names
    Set dicFKs = getFKs(pTdf, pSelectedDB)
    If dicFKs Is Nothing Then GoTo ExitHandler
    
    For i = 1 To pTdf.Fields.Count
        Set fld = pTdf.Fields(i - 1)
        'Field Name
        defArr(i, 0) = fld.Name
        'Data Type
        defArr(i, 1) = getFieldTypeString(fld.Type)
        'Size
        If fld.Type = dbText Then
            defArr(i, 2) = fld.Size
        Else
            defArr(i, 2) = "-"
        End If
        'Required or not
        If fld.Required Then defArr(i, 3) = ChrW("&H" & 2714)
        'Primary key or not
        If dicPKs.Exists(fld.Name) Then defArr(i, 4) = ChrW("&H" & 2714)
        'Foreign key or not
        If dicFKs.Exists(fld.Name) Then defArr(i, 5) = ChrW("&H" & 2714)
        'Description
        On Error Resume Next
        description = fld.Properties("Description")
        On Error GoTo ErrorHandler
        defArr(i, 6) = description
    Next i


    getTableDefArray = defArr
    
ExitHandler:
    
    Set fld = Nothing
    Set dicFKs = Nothing
    Set dicPKs = Nothing
    
    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.description, vbCritical, Tool_Name
        
    GoTo ExitHandler
        
End Function


'******************************************************************************************
'*Function      : get field type string of argument field
'*arg(1)        : field type number
'*return        : field data string

'******************************************************************************************
Public Function getFieldTypeString(ByVal pFldTyepNum As Long) As String
    
    'Const
    Const FUNC_NAME As String = "getFieldTypeString"
    
    'Variable
    Dim strType As String
    
    On Error GoTo ErrorHandler

    strType = ""
    
    Select Case pFldTyepNum
    Case dbBoolean
        strType = "Bool"
    Case dbByte
        strType = "Byte"
    Case dbInteger
        strType = "Integer"
    Case dbLong
        strType = "Long Integer"
    Case dbSingle
        strType = "Single Number"
    Case dbDouble
        strType = "Double Number"
    Case dbCurrency
        strType = "Currency"
    Case dbDate
        strType = "Date"
    Case dbText
        strType = "short Text"
    Case dbLongBinary
        strType = "OLE Object Type"
    Case dbMemo
        strType = "long text"
    End Select


    getFieldTypeString = strType
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.description, vbCritical, Tool_Name
        
    GoTo ExitHandler
        
End Function


'******************************************************************************************
'*Function      : get field strings of primary keys
'*arg(1)        : TableDef instance
'*return        : Dictionary containing PK info
'******************************************************************************************
Public Function getPKs(ByVal pTdf As DAO.TableDef) As Object
    
    'Const
    Const FUNC_NAME As String = "getPKs"
    
    'Variable
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    Dim dic As Object
    
    On Error GoTo ErrorHandler

    Set getPKs = Nothing
    Set dic = CreateObject("Scripting.Dictionary")
    
    
    'check if Primary property is true
    For Each idx In pTdf.Indexes
        If idx.Primary = True Then
            For Each fld In idx.Fields
                dic.Add fld.Name, True
            Next
        End If
    Next

    'Return
    Set getPKs = dic
    
ExitHandler:

    Set dic = Nothing

    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.description, vbCritical, Tool_Name
        
    GoTo ExitHandler
        
End Function


'******************************************************************************************
'*Function      : get field strings of foreign keys
'*arg(1)        : TableDef instance
'*arg(2)        : DAO Database instance
'*return        : Dictionary containing PK info

'******************************************************************************************
Public Function getFKs( _
       ByVal pTdf As DAO.TableDef, _
       ByVal pSelectedDB As DAO.Database _
       ) As Object
    
    'Const
    Const FUNC_NAME As String = "getFKs"
    
    'Variable
    Dim rsRelation As DAO.Recordset
    Dim dic As Object
    
    On Error GoTo ErrorHandler

    Set getFKs = Nothing
    Set dic = CreateObject("Scripting.Dictionary")
    
    'access MSysRelationships system table
    Set rsRelation = pSelectedDB.OpenRecordset( _
                     "SELECT szColumn FROM MSysRelationships WHERE szObject =" & _
                     " " & _
                     "'" & _
                     pTdf.Name & _
                     "'" & _
                     ";" _
                     )
    
    With rsRelation
        If .EOF Then Set getFKs = dic: GoTo ExitHandler
        .MoveFirst
        Do Until .EOF
            dic.Add .Fields("szColumn").Value, True
            .MoveNext
        Loop
    End With
    
    'Return
    Set getFKs = dic
    
ExitHandler:
    
    If Not rsRelation Is Nothing Then rsRelation.Close: Set rsRelation = Nothing
        
    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.description, vbCritical, Tool_Name

        
    GoTo ExitHandler
        
End Function


'******************************************************************************************
'*Function      : set worksheet Name to argument sheet
'*arg(1)        : excel worksheet instance
'*arg(2)        : the name set
'*return        : True > normal termination; False > abnormal termination

'******************************************************************************************
Public Function setWSName( _
       ByVal ws As Object, _
       ByVal newName As String _
       ) As Boolean
    
    'Const
    Const FUNC_NAME As String = "setWSName"
    
    'Variable
    
    On Error GoTo ErrorHandler

    setWSName = False
    
    ws.Name = newName

    setWSName = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    'escaping route: if the name includes some charactors not allowed to use to sheet name
    ws.Name = "Table_" & ws.Parent.Worksheets.Count & "_" & Format(Now, "yyyymmddhhnnss")

    setWSName = True
    GoTo ExitHandler
        
End Function



