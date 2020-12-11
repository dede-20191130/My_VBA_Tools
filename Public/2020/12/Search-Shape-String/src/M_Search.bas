Attribute VB_Name = "M_Search"
'@Folder("VBAProject")
Option Explicit
'**************************
'*�}�`�̕����񌟍��E�u��
'*
'*referencing https://qiita.com/s-hchika/items/dda585fa0bdb829e9713
'**************************

'�萔
'�|�b�v�A�b�v�̖��O
Private Const TITLE_SEARCH_SHAPE_TEXT As String = "�I�[�g�V�F�C�v����"

'�ϐ�



'******************************************************************************************
'*�֐���    �F���������֐�
'*�@�\      �F
'*����      �F
'******************************************************************************************
Public Sub searchShapeText()

    
    '�萔
    Const FUNC_NAME As String = "searchShapeText"
    
    '�ϐ�
    Dim mySheets As Variant                     '���[�N�V�[�g�̏W����
    Dim sheet As Variant
    Dim searchWord As String                     '�������[�h
    Dim flgTerminate As Boolean
    Dim flgFound As Boolean
    
    On Error GoTo ErrorHandler
    
    '�u�b�N������or�V�[�g����
    If MsgBox("�u�b�N�S�̂������ꏊ�Ƃ��܂����B", vbYesNo, TITLE_SEARCH_SHAPE_TEXT) = vbYes Then
        '�Ώۂ̃��[�N�V�[�g�����݊J���Ă���u�b�N�̑S�ẴV�[�g�Ƃ���
        Set mySheets = ActiveWorkbook.Worksheets
    Else
        '�Ώۂ̃��[�N�V�[�g�����݊J���Ă���V�[�g�݂̂Ƃ���
        mySheets = Array(ActiveSheet)
    End If
    
    '�������[�h���̓|�b�v�A�b�v��\������
    searchWord = Trim(InputBox("�������������[�h����͂��ĉ������B", TITLE_SEARCH_SHAPE_TEXT))

    If searchWord = "" Then GoTo ExitHandler
    
    '����
    For Each sheet In mySheets
        sheet.Activate
        If Not searchReplaceShapeText(sheet.Shapes, searchWord, flgTerminate, flgFound) Then GoTo ExitHandler
        '�I���t���OTrue�̏ꍇ
        If flgTerminate Then GoTo ExitHandler
    Next sheet
    
    '���ׂĂ̌����͈͂Ŗ������̏ꍇ
    If Not flgFound Then MsgBox "�u" & searchWord & "�v��������܂���B", vbExclamation, TITLE_SEARCH_SHAPE_TEXT
    
ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TITLE_SEARCH_SHAPE_TEXT
        
    GoTo ExitHandler
        
End Sub


'******************************************************************************************
'*�֐���    �FsearchReplaceShapeText
'*�@�\      �F�}�`�������u���֐�
'*����      �FworksheetShapes Worksheet�̐}�`�R���N�V����
'*����      �FsearchWord      ��������
'*����      �FflgTerminate      �T���I���t���O
'*����      �FflgFound      �����񔭌��t���O
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
Private Function searchReplaceShapeText(ByVal worksheetShapes As Object, ByVal searchWord As String, _
                                        ByRef flgTerminate As Boolean, ByRef flgFound As Boolean) As Boolean

    
    '�萔
    Const FUNC_NAME As String = "searchReplaceShapeText"
    
    '�ϐ�
    Dim targetShape  As Excel.Shape              '���[�N�V�[�g���̐}�`
    Dim shapeText   As String                    '�}�`���̕���
    Dim discoveryWord As Long                    '�������[�h�����ʒu
    Dim replaceWord As String                    '�u����̕���
    Dim replacePopupMsg As String                '�u���|�b�v�A�b�v���b�Z�[�W
    Dim searchWordCnt As Long: searchWordCnt = 1 '�}�`���������[�h��
    
    On Error GoTo ErrorHandler


    '���[�N�V�[�g�ɐ}�`�����݂���ԃ��[�v
    For Each targetShape In worksheetShapes
        Do

            '�N���[�v�����ꂽ�}�`�̎�
            If (targetShape.Type = msoGroup) Then
    
                If Not (searchReplaceShapeText(targetShape.GroupItems, searchWord, flgTerminate, flgFound)) Then GoTo ExitHandler
                '�I���t���OTrue�̏ꍇ
                If flgTerminate Then GoTo TruePoint
    
                '�R�����g�̎�
            ElseIf (targetShape.Type = msoComment) Then
                Exit Do
            Else
                '�w�肵���e�L�X�g�t���[���Ƀe�L�X�g�����邩�ǂ�����Ԃ�
                If (targetShape.TextFrame2.HasText) Then
    
                    '�}�`���̃e�L�X�g���擾
                    shapeText = targetShape.TextFrame2.TextRange.Text
    
                    '�}�`���̕����񂩂猟��
                    discoveryWord = InStr(shapeText, searchWord)
    
                    '�������[�h�����������Ƃ��A�u���̏������s��
                    If (discoveryWord > 0&) Then
                        
                        '�����񔭌��t���OTrue
                        flgFound = True
                        
                        '�E�B���h�E��}�`�̈ʒu�ɃX�N���[��
                        ActiveWindow.ScrollRow = targetShape.TopLeftCell.Row
                        ActiveWindow.ScrollColumn = targetShape.TopLeftCell.Column
    
                        Do While (discoveryWord > 0&)
                            
                            '�e�L�X�g�͈͑I�����������邽�߁A�J�����g�Z����I������
                            targetShape.TopLeftCell.Select
    
                            targetShape.TextFrame2.TextRange.Characters(discoveryWord, Len(searchWord)).Select
    
                            replacePopupMsg = "�u������ꍇ�A���͂��Ă��������B" & vbNewLine & vbNewLine & "�u���O : " & searchWord & vbNewLine & "�u����"
    
                            ' �u�����̓��b�Z�[�W���o�͂���
                            replaceWord = InputBox(replacePopupMsg, "�u��")
    
                            If Not replaceWord = "" Then
                            
                                '�}�`���̕��������ӏ��u������
                                targetShape.TextFrame2.TextRange.Text = Replace(shapeText, searchWord, replaceWord, 1, searchWordCnt)
                                targetShape.TopLeftCell.Select
    
                            End If
    
                            '�������p�����邩�ǂ���
                            If (MsgBox("continue?", vbQuestion Or vbOKCancel, TITLE_SEARCH_SHAPE_TEXT) <> vbOK) Then
                                flgTerminate = True
                                GoTo TruePoint
    
                                '�����}�`���ŕ�������
                            Else
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

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TITLE_SEARCH_SHAPE_TEXT
        
    GoTo ExitHandler
        
End Function



