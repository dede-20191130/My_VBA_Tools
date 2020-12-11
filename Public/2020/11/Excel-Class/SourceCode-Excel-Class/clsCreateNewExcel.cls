VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCreateNewExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("class")
Option Compare Database
Option Explicit

'**************************
'*Excel�u�b�N�쐬�N���X
'**************************

'�萔
Private currentSheetNum As Long

'�ϐ�
Private myXlApp As Object
Private myWorkBook As Object
Private dicWorkSheet As Dictionary '�u�b�N�̑S�Ẵ��[�N�V�[�g���i�[����Ă���Dictionary�I�u�W�F�N�g


'******************************************************************************************
'*getter/setter
'******************************************************************************************
Public Property Get xlApplication() As Object
    Set xlApplication = myXlApp
End Property


Public Property Get Workbook() As Object
    Set Workbook = myWorkBook
End Property


Public Property Get WorkSheets(ByVal num As Long) As Object
    If Not dicWorkSheet.Exists(num) Then Call MsgBox("�V�[�g������܂���B", vbExclamation, TOOL_NAME): Set WorkSheets = Nothing: Exit Property
    Set WorkSheets = dicWorkSheet.Item(num)
End Property


'******************************************************************************************
'*�֐���    �FClass_Initialize
'*�@�\      �F
'*����      �F
'******************************************************************************************
Private Sub Class_Initialize()
    
    '�萔
    Const FUNC_NAME As String = "Class_Initialize"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler
    
    '�V�[�g���̐ݒ�
    currentSheetNum = 1
    
    '�I�u�W�F�N�g�̐ݒ�
    Set myXlApp = CreateObject("Excel.Application")
    With myXlApp
        '�o�b�N�O���E���h�Ő����E�ҏW
        .Visible = False
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With
    Set myWorkBook = myXlApp.Workbooks.Add
    Set dicWorkSheet = New Dictionary
    dicWorkSheet.Add currentSheetNum, myWorkBook.WorkSheets(currentSheetNum)
    
ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Sub


'******************************************************************************************
'*�֐���    �FClass_Terminate
'*�@�\      �F
'*����      �F
'******************************************************************************************
Private Sub Class_Terminate()
    
    '�萔
    Const FUNC_NAME As String = "Class_Terminate"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler
    
    '�u�b�N�̕ۑ� To:=�f�X�N�g�b�v
    With CreateObject("WScript.Shell")
        myWorkBook.SaveAs .SpecialFolders("Desktop") & "\" & "Test-Excel-Class-" & Format(Now, "yyyymmddhhnnss") & ".xlsx"
    End With
    
    '�ݒ�̃��X�g�A
    With myXlApp
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With
    
    '�N���[�Y
    myWorkBook.Close
    myXlApp.Quit

ExitHandler:
    
    Set dicWorkSheet = Nothing
    Set myWorkBook = Nothing
    Set myXlApp = Nothing
    
    Exit Sub
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Sub



'******************************************************************************************
'*�֐���    �FaddNewSheet
'*�@�\      �F�u�b�N�ɃV�[�g��ǉ�
'*����      �F
'*�߂�l    �F�ǉ����ꂽ�V�[�g�I�u�W�F�N�g
'******************************************************************************************
Public Function addNewSheet() As Object
    
    '�萔
    Const FUNC_NAME As String = "addNewSheet"
    
    '�ϐ�
    Dim ws As Object
    
    On Error GoTo ErrorHandler

    Set addNewSheet = Nothing
    
    currentSheetNum = currentSheetNum + 1
    '�Ō���ɃV�[�g�ǉ�
    Set ws = myWorkBook.WorkSheets.Add(After:=myWorkBook.WorkSheets(myWorkBook.WorkSheets.Count))
    dicWorkSheet.Add currentSheetNum, ws
    
    Set addNewSheet = ws
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Function
