Attribute VB_Name = "mdlChangePrintArea"
Option Explicit

'**************************
'*PageSetup_PrintArea�ύX�e�X�g
'**************************

'�萔��
Private Const SOURCE_NAME As String = "mdlChangePrintArea"


'�ϐ���



'******************************************************************************************
'*getter/setter��
'******************************************************************************************





'******************************************************************************************
'*�֐���    �FchangePrintAreaBeforeRevised
'*�@�\      �FPrintArea���ЂƂ��̍s�ɕύX���� �C���O
'*����      �F
'******************************************************************************************
Public Sub changePrintAreaBeforeRevised()
    
    '�萔
    Const FUNC_NAME As String = "changePrintAreaBeforeRevised"
    
    '�ϐ�
    Dim prePrintAreaAddress As String
    
    On Error GoTo ErrorHandler
    
    With ThisWorkbook.Worksheets(1)
    
        '���݂̈���͈̓A�h���X
        prePrintAreaAddress = .PageSetup.PrintArea
        
        '����͈͂��ЂƂ��̍s�ɕύX����
        .PageSetup.PrintArea = .Range(prePrintAreaAddress).Resize(.Range(prePrintAreaAddress).Rows.Count + 1).Address
        
        Debug.Print .PageSetup.PrintArea
        
    End With

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, SOURCE_NAME
        
    GoTo ExitHandler
        
End Sub





'******************************************************************************************
'*�֐���    �FchangePrintAreaBeforeRevised
'*�@�\      �FPrintArea���ЂƂ��̍s�ɕύX���� �C��01
'*����      �F
'******************************************************************************************
Public Sub changePrintAreaRevised01()
    
    '�萔
    Const FUNC_NAME As String = "changePrintAreaRevised01"
    
    '�ϐ�
    Dim prePrintAreaAddress As String
    Dim currentStyle As XlReferenceStyle

    On Error GoTo ErrorHandler
    
    With ThisWorkbook.Worksheets(1)
        
        '�Z���̎Q�ƌ`����A1�`���ɕύX
        currentStyle = Application.ReferenceStyle
        Application.ReferenceStyle = xlA1
        
        '���݂̈���͈̓A�h���X
        prePrintAreaAddress = .PageSetup.PrintArea
        
        '����͈͂��ЂƂ��̍s�ɕύX����
        .PageSetup.PrintArea = .Range(prePrintAreaAddress).Resize(.Range(prePrintAreaAddress).Rows.Count + 1).Address
        
        Debug.Print .PageSetup.PrintArea
        
        '�Z���̎Q�ƌ`���𕜋�����
        Application.ReferenceStyle = currentStyle
        
    End With

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, SOURCE_NAME
        
    GoTo ExitHandler
        
End Sub




'******************************************************************************************
'*�֐���    �FchangePrintAreaBeforeRevised
'*�@�\      �FPrintArea���ЂƂ��̍s�ɕύX���� �C��02
'*����      �F
'******************************************************************************************
Public Sub changePrintAreaRevised02()
    
    '�萔
    Const FUNC_NAME As String = "changePrintAreaRevised02"
    
    '�ϐ�
    Dim prePrintAreaAddress As String
    
    On Error GoTo ErrorHandler
    
    With ThisWorkbook.Worksheets(1)
    
        '���݂̈���͈̓A�h���X
        prePrintAreaAddress = .PageSetup.PrintArea
        
        '�A�h���X��xlA1�Q�ƌ`���̂��̂ɏC��
        If Application.ReferenceStyle = xlR1C1 Then prePrintAreaAddress = Application.ConvertFormula(prePrintAreaAddress, xlR1C1, xlA1)
        
        '����͈͂��ЂƂ��̍s�ɕύX����
        .PageSetup.PrintArea = .Range(prePrintAreaAddress).Resize(.Range(prePrintAreaAddress).Rows.Count + 1).Address
        
        Debug.Print .PageSetup.PrintArea
        
    End With

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, SOURCE_NAME
        
    GoTo ExitHandler
        
End Sub

