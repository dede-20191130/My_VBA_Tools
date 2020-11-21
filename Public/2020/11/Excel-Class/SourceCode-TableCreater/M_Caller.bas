Attribute VB_Name = "M_Caller"
'@Folder("Module")
Option Explicit

'**************************
'*TableCreater�N���X�̌Ăяo����
'**************************

'�萔

'�ϐ�



'******************************************************************************************
'*getter/setter
'******************************************************************************************


'******************************************************************************************
'*�֐���    �FTestTemplateA
'*�@�\      �F���{�̃e���v��A�ɂ��āATableCreater��p���ĕ\���쐬����
'               �쐬�ꏊ�F�V�K�V�[�g
'*����      �F
'******************************************************************************************
Public Sub TestTemplateA()
    
    '�萔
    Const FUNC_NAME As String = "TestTemplateA"
    
    '�ϐ�
    Dim ws As Worksheet
    Dim objTableCreater As tableCreater
    
    On Error GoTo ErrorHandler
    
    With ThisWorkbook
        '�V�K�V�[�g�쐬
        Set ws = .Worksheets.Add(, .Worksheets(.Worksheets.Count))
        ws.Name = "�e���v��A_" & Format(Now, "yyyymmddhhnnss")
        
        '���{���e���v�����R�s�[
        ws.Range(ws.Cells(2, 2), ws.Cells(9, 4)).Value = .Worksheets("���{").Range(.Worksheets("���{").Cells(2, 2), .Worksheets("���{").Cells(9, 4)).Value
        
        'TableCreater���I�u�W�F�N�g��
        Set objTableCreater = New tableCreater
        '�͈͂Ə��v���ݒ�
        Set objTableCreater.Range = ws.Range(ws.Cells(2, 2), ws.Cells(9, 4))
        objTableCreater.ColumnSubTotal = 4
        
        '�r�������� �ُ�I������ExitHandler�i�I�������j�Ɉڍs
        If Not objTableCreater.drawLines Then GoTo ExitHandler
         
        '�w�b�_�[�̋����̂��߂̃X�^�C���ύX���s�� �ُ�I������ExitHandler�i�I�������j�Ɉڍs
        If Not objTableCreater.setStyleForHeader Then GoTo ExitHandler
        
        '���v���獇�v���v�Z �ُ�I������ExitHandler�i�I�������j�Ɉڍs
        If Not objTableCreater.calcTotalFromSubTotal Then GoTo ExitHandler
        
        '�s���E�񕝂̒���
        ws.Range(ws.Cells(2, 2), ws.Cells(9, 4)).EntireColumn.AutoFit
        
    End With
    

ExitHandler:
    
    '�ϐ������
    Set objTableCreater = Nothing
    Set ws = Nothing
    
    Exit Sub
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "TableCreater"
        
    GoTo ExitHandler
        
End Sub






'******************************************************************************************
'*�֐���    �FTestTemplateB
'*�@�\      �F���{�̃e���v��B�ɂ��āATableCreater��p���ĕ\���쐬����
'               �쐬�ꏊ�F�V�K�V�[�g
'*����      �F
'******************************************************************************************
Public Sub TestTemplateB()
    
    '�萔
    Const FUNC_NAME As String = "TestTemplateB"
    
    '�ϐ�
    Dim ws As Worksheet
    Dim objTableCreater As tableCreater
    
    On Error GoTo ErrorHandler
    
    With ThisWorkbook
        '�V�K�V�[�g�쐬
        Set ws = .Worksheets.Add(, .Worksheets(.Worksheets.Count))
        ws.Name = "�e���v��B_" & Format(Now, "yyyymmddhhnnss")
        
        '���{���e���v�����R�s�[
        ws.Range(ws.Cells(2, 2), ws.Cells(8, 10)).Value = .Worksheets("���{").Range(.Worksheets("���{").Cells(12, 2), .Worksheets("���{").Cells(18, 10)).Value
        
        'TableCreater���I�u�W�F�N�g��
        Set objTableCreater = New tableCreater
        '�͈͂Ə��v���ݒ�
        Set objTableCreater.Range = ws.Range(ws.Cells(2, 2), ws.Cells(8, 10))
        objTableCreater.ColumnSubTotal = 10
        
        '�r�������� �ُ�I������ExitHandler�i�I�������j�Ɉڍs
        If Not objTableCreater.drawLines Then GoTo ExitHandler
         
        '�w�b�_�[�̋����̂��߂̃X�^�C���ύX���s�� �ُ�I������ExitHandler�i�I�������j�Ɉڍs
        If Not objTableCreater.setStyleForHeader Then GoTo ExitHandler
        
        '���v���獇�v���v�Z �ُ�I������ExitHandler�i�I�������j�Ɉڍs
        If Not objTableCreater.calcTotalFromSubTotal Then GoTo ExitHandler
        
        '�s���E�񕝂̒���
        ws.Range(ws.Cells(2, 2), ws.Cells(8, 10)).EntireColumn.AutoFit
        
    End With
    

ExitHandler:
    
    '�ϐ������
    Set objTableCreater = Nothing
    Set ws = Nothing
    
    Exit Sub
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "TableCreater"
        
    GoTo ExitHandler
        
End Sub

