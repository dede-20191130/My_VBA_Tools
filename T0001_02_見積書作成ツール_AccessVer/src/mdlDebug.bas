Attribute VB_Name = "mdlDebug"
'@Folder("Database")
Option Compare Database
Option Explicit

#If DEBUG_MODE Then

Sub s20200515_1137()

    Dim cnt As Variant
    
    For Each cnt In CurrentDb.TableDefs
        Debug.Print cnt.Name
    Next
    
    
End Sub


Sub s20200517_1939()
    Debug.Print CurrentData.AllTables(12).Name
End Sub


Private Sub AccessObjectTest20200517_1945()

    Dim obj As AccessObject

    'DataAccessPageを列挙
    Debug.Print "DataAccessPageを列挙"
    For Each obj In CurrentProject.AllDataAccessPages
        Debug.Print obj.Name
    Next

    'フォームを列挙
    Debug.Print "フォームを列挙"
    For Each obj In CurrentProject.AllForms
        Debug.Print obj.Name
    Next

    'マクロを列挙
    Debug.Print "マクロを列挙"
    For Each obj In CurrentProject.AllMacros
        Debug.Print obj.Name
    Next

    'モジュールを列挙
    Debug.Print "モジュールを列挙"
    For Each obj In CurrentProject.AllModules
        Debug.Print obj.Name
    Next

    'レポートを列挙
    Debug.Print "レポートを列挙"
    For Each obj In CurrentProject.AllReports
        Debug.Print obj.Name
    Next

    'テーブルの列挙
    Debug.Print "テーブルの列挙"
    For Each obj In CurrentData.AllTables
        Debug.Print obj.Name
    Next

    'クエリの列挙
    Debug.Print "クエリの列挙"
    For Each obj In CurrentData.AllQueries
        Debug.Print obj.Name
    Next

End Sub


'Sub s20201017_2327()
'    #If Not CBool(DEBUG_MODE) Then
'        Exit Sub
'    #End If
'
'    Dim c
'
'    With CurrentDb.OpenRecordsexxxt("select * from 見積詳細データ order by 見積番号")
'        .MoveFirst
'        For Each c In .Fields
'            Debug.Print c.value
'        Next
'        .MoveNext
'        For Each c In .Fields
'            Debug.Print c.value
'        Next
'        .MoveNext
'        For Each c In .Fields
'            Debug.Print c.value
'        Next
'        .MoveNext
'        For Each c In .Fields
'            Debug.Print c.value
'        Next
'    End With
'
'End Sub


Sub s20201023_1558()
    Call f20201023_1557(TBL_T_ESTM_DTL)
End Sub


Function f20201023_1557(TableName As String)
    Dim db As dao.Database
    Dim tbd As dao.TableDef
    Dim idx As dao.Index
    Dim fld As dao.Field
    Set db = CurrentDb
    Set tbd = db.TableDefs(TableName)
    For Each idx In tbd.Indexes
        If idx.Primary = True Then
            For Each fld In idx.Fields
                Debug.Print fld.Name
            Next
        End If
    Next
End Function





Sub s20201109_2316()
    Dim a, b, c, x
    Dim d As Boolean
    a = "a"
    b = "b"
    c = "c"
    x = 123
    
    debugPrintDev (checkWhetherControlsVacant(d, a, b, c))
    debugPrintDev (d)
    
End Sub


Sub s20201109_2345()
    Dim a As Long
    Dim b As Double
    Dim c As String
    a = 20000
    c = "1000000000000000000000000000000000.0000000001"
    
    b = a
    b = c
    
End Sub

Sub s20210123_2033()
    Dim mc As MatchCollection
    Dim m As Match
    Dim objReg As New clsWrappedRegExp
    Dim v As Variant
    Set mc = objReg.execute("Abc123DEFGH4567ijkl890", "(\D+)(\d+)")
    For Each m In mc
        With m
            Debug.Print .FirstIndex
            Debug.Print .VALUE
            For Each v In .SubMatches
                Debug.Print v
            Next
        End With
    Next
End Sub

Sub s20210123_2046()
    Dim objReg As New clsWrappedRegExp
    Debug.Print objReg.replace("Abc123あいう4567def890", "([A-ZＡ-Ｚ]+)([0-9]+)", "$1")
    Debug.Print objReg.replace("Abc123あいう4567def890", "([A-ZＡ-Ｚ]+)([0-9]+)", "$2")
End Sub

Sub s20210123_2053()
    Dim objReg As New clsWrappedRegExp
    Dim s
    s = "^[\d-]+$"
    Debug.Print objReg.test("08022239920", s)
    Debug.Print objReg.test("080-2223-9920", s)
    Debug.Print objReg.test("080222399202323434343---------", s)
    Debug.Print objReg.test("080_2223_9920", s)
    Debug.Print objReg.test("080a2223a9920", s)
    Debug.Print objReg.test("234", "[^\d]")
End Sub



Sub s20210210_0028()
    Dim coll As New Collection
    Dim v As Variant
    
    coll.Add "a"
    coll.Add "b"
    coll.Add "c"
    
    v = CollectionToArray(coll)
    
    Debug.Print Join(CollectionToArray(coll), ";")
    Debug.Print Join(CollectionToArray(coll), vbNewLine)
    
    
End Sub


Sub s20210220_2334()
    With CurrentDb.OpenRecordset(QRY_Q_MAX_W_ESTM)
        Debug.Print "a" & .RecordCount & "b"
        Debug.Print .Fields(0).VALUE
        
    End With
    
End Sub

Sub s20210223_2150()
    Dim o As New clsManageExcelBook
    Dim wb
    Set wb = o.addExistingBook("C:\temp\20210223\a.xlsx")
    
    Debug.Print o.WorkSheets(1).Name
    
    o.addNewSheet
    o.WorkSheets(2).Name = "orange"
    
    
End Sub

Sub s20210227_0000()
    With CurrentDb.OpenRecordset("tmp01")
        .MoveFirst
        Dim c
        For Each c In .Fields
            Debug.Print c.VALUE
        Next
    End With
End Sub

#End If
