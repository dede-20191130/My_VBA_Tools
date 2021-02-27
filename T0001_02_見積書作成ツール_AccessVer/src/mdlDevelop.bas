Attribute VB_Name = "mdlDevelop"
'@Folder("Database")
Option Compare Database
Option Explicit

#If DEBUG_MODE Then



Sub debugPrintDev(x As Variant)
    #If Not CBool(DEBUG_MODE) Then
        MsgBox "DEBUG_MODEではないためこのコードを削除してください"
        Exit Sub
    #End If
    Debug.Print Now & " " & CStr(x)
End Sub



'******************************************************************************************
'*関数名    ：exportCodesSQLs
'*機能      ：モジュール・クラスのコード及びクエリのSQLの出力
'*引数      ：
'******************************************************************************************
Sub exportCodesSQLs()
    
    '定数
    Const FUNC_NAME As String = "exportCodesSQLs"
    
    '変数
    Dim outputDir As String
    Dim vbcmp As Object
    Dim fileName As String
    Dim ext As String
    Dim qry As QueryDef
    Dim qName As String
    
    
    
    On Error GoTo ErrorHandler
    
    outputDir = _
        Access.CurrentProject.Path & _
        "\" & _
        "src_" & _
        Left(Access.CurrentProject.Name, InStrRev(Access.CurrentProject.Name, ".") - 1)
    If Dir(outputDir) = "" Then MkDir outputDir
    
    'モジュール・クラスの出力
    For Each vbcmp In VBE.ActiveVBProject.VBComponents
        With vbcmp
            '拡張子
            Select Case .Type
            Case 1
                ext = ".bas"
            Case 2, 100
                ext = ".cls"
            Case 3
                ext = ".frm"
            End Select
                        
            fileName = .Name & ext
            fileName = gainStrNameSafe(fileName) 'ファイル名に使用できない文字を置換
            If fileName = "" Then GoTo ExitHandler
            
            'output
            .Export outputDir & "\" & fileName
            
        End With
    Next vbcmp
    
    'SQLの出力
    With CreateObject("Scripting.FileSystemObject")
        For Each qry In CurrentDb.QueryDefs
            Do
                qName = gainStrNameSafe(qry.Name) 'ファイル名に使用できない文字を置換
                If qName = "" Then GoTo ExitHandler
                
                If qName Like "Msys*" Then Exit Do 'システム関連クエリは除外
                
                With .CreateTextFile(outputDir & "\" & qName & ".sql")
                    .write qry.SQL
                    .Close
                End With
            Loop While False
        Next qry
    End With

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & err.Number & vbNewLine & _
           err.description, vbCritical, "マクロ"
        
    GoTo ExitHandler
        
End Sub




'******************************************************************************************
'*関数名    ：gainStrNameSafe
'*機能      ：ファイル名に使用できない文字をアンダースコアに置換する
'*引数      ：対象の文字列
'*戻り値    ：置換後文字列
'******************************************************************************************
Public Function gainStrNameSafe(ByVal s As String) As String
    
    '定数
    Const FUNC_NAME As String = "gainStrNameSafe"
    
    '変数
    Dim x As Variant
    
    On Error GoTo ErrorHandler

    gainStrNameSafe = ""
    
    For Each x In Split("\,/,:,*,?,"",<,>,|", ",") 'ファイル名に使用できない文字の配列
        s = replace(s, x, "_")
    Next x
    
    gainStrNameSafe = s

ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & err.Number & vbNewLine & _
           err.description, vbCritical, "マクロ"
        
    GoTo ExitHandler
        
End Function




#End If
