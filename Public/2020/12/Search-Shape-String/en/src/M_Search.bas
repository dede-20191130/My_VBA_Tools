Attribute VB_Name = "M_Search"
'@Folder("VBAProject")
Option Explicit
'**************************
'*search and replace a text of shape
'*
'*referencing https://qiita.com/s-hchika/items/dda585fa0bdb829e9713
'**************************

'Consts
'popup title
Private Const TITLE_SEARCH_SHAPE_TEXT As String = "Auto Shape Search"

'Vars
'None


'******************************************************************************************
'*Function :the main processing of searching function
'******************************************************************************************
Public Sub searchMain()

    
    'Consts
    Const FUNC_NAME As String = "searchMain"
    
    'Vars
    Dim mySheets As Variant                     'collection of worksheets
    Dim sheet As Variant
    Dim searchWord As String
    Dim flgTerminate As Boolean
    Dim flgFound As Boolean
    
    On Error GoTo ErrorHandler
    
    'search through the entire book or search through one sheet
    If MsgBox("Do you want to search through the entire book?", vbYesNo, TITLE_SEARCH_SHAPE_TEXT) = vbYes Then
        'the target is all sheets of current open book
        Set mySheets = ActiveWorkbook.Worksheets
    Else
        'the target is only a active sheet
        mySheets = Array(ActiveSheet)
    End If
    
    'display a popup window to input the searched word
    searchWord = Trim(InputBox("Input the word you want to search.", TITLE_SEARCH_SHAPE_TEXT))

    If searchWord = "" Then GoTo ExitHandler
    
    'perform a search
    For Each sheet In mySheets
        sheet.Activate
        If Not searchReplaceShapeText(sheet.Shapes, searchWord, flgTerminate, flgFound) Then GoTo ExitHandler
        If flgTerminate Then GoTo ExitHandler
    Next sheet
    
    'if not found, message is shown
    If Not flgFound Then MsgBox """" & searchWord & """ is not found.", vbExclamation, TITLE_SEARCH_SHAPE_TEXT
    
ExitHandler:

    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TITLE_SEARCH_SHAPE_TEXT
        
    GoTo ExitHandler
        
End Sub




'******************************************************************************************
'*Function :searching and replacing the text of the shape in target shape-collection
'*Arg      :shape-collection in the worksheet
'*Arg      :word for search
'*Arg      :termination flag
'*Arg      :flag of having found the word
'*Retrun    :True > normal termination; False > abnormal termination
'******************************************************************************************
Private Function searchReplaceShapeText(ByVal worksheetShapes As Object, ByVal searchWord As String, _
                                        ByRef flgTerminate As Boolean, ByRef flgFound As Boolean) As Boolean

    
    'Consts
    Const FUNC_NAME As String = "searchReplaceShapeText"
    
    'Vars
    Dim targetShape  As Excel.Shape              'current target shape
    Dim shapeText   As String                    'text of the shape
    Dim discoveryWord As Long                    'the position in which the word is discovered
    Dim replaceWord As String                    'word after replacing
    Dim replacePopupMsg As String                'popup message for replacing
    Dim searchWordCnt As Long: searchWordCnt = 1 'word count in the shape
    
    On Error GoTo ErrorHandler


    For Each targetShape In worksheetShapes
        Do

            If (targetShape.Type = msoGroup) Then
                'if target shape is grouped

                'call itself recursively
                If Not (searchReplaceShapeText(targetShape.GroupItems, searchWord, flgTerminate, flgFound)) Then GoTo ExitHandler
                'exit if termination flagged
                If flgTerminate Then GoTo TruePoint
    
            ElseIf (targetShape.Type = msoComment) Then
                'continue if it's comment object
                Exit Do
            Else
                'check if it has text
                If (targetShape.TextFrame2.HasText) Then
    
                    'get the text
                    shapeText = targetShape.TextFrame2.TextRange.Text
    
                    'get the position of hit word
                    discoveryWord = InStr(shapeText, searchWord)
    
                    'process replacing block if discovered
                    If (discoveryWord > 0&) Then
                        
                        'found flagged
                        flgFound = True
                        
                        'scroll to the position of the shape
                        ActiveWindow.ScrollRow = targetShape.TopLeftCell.Row
                        ActiveWindow.ScrollColumn = targetShape.TopLeftCell.Column
    
                        Do While (discoveryWord > 0&)
                            
                            'select current cell to cancel the previous selection of text range
                            targetShape.TopLeftCell.Select

                            'select target text
                            targetShape.TextFrame2.TextRange.Characters(discoveryWord, Len(searchWord)).Select

                            replacePopupMsg = "Input any text if you want to replace it with." & vbNewLine & vbNewLine & "Before: " & searchWord & vbNewLine & "After: "
    
                            'show inquiry message
                            replaceWord = InputBox(replacePopupMsg, "Replace")
    
                            If Not replaceWord = "" Then
                            
                                'replace a hit text with given text
                                targetShape.TextFrame2.TextRange.Text = Replace(shapeText, searchWord, replaceWord, 1, searchWordCnt)
                                targetShape.TopLeftCell.Select
    
                            End If
    
                            'inquire if continue
                            If (MsgBox("continue?", vbQuestion Or vbOKCancel, TITLE_SEARCH_SHAPE_TEXT) <> vbOK) Then
                                flgTerminate = True
                                GoTo TruePoint
    
                            Else
                                'search text in the same shape
                                discoveryWord = InStr(discoveryWord + 1&, shapeText, searchWord)
                            End If
    
                        Loop
    
                    End If
                End If
            End If
        Loop While False
    Next
    

TruePoint:

    searchReplaceShapeText = True

ExitHandler:
    
    
    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TITLE_SEARCH_SHAPE_TEXT
        
    GoTo ExitHandler
        
End Function

