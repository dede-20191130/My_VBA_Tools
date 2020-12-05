Attribute VB_Name = "M_Test"
Option Explicit



'**************************
'*��O�����T���v��
'**************************






'******************************************************************************************
'*�֐���    �F��O����Sub�v���V�[�W������
'*�@�\      �F
'*����      �F
'******************************************************************************************
Public Sub subSample()
    
    '�萔
    Const FUNC_NAME As String = "subSample"
    
    '�ϐ�
    Dim filePathArr As Variant
    Dim filePath As Variant
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    'funcSample01�̌Ăяo���@���݂��Ȃ��V�[�g���̈����ŌĂяo��
    filePathArr = funcSample01("sheetNotExist")
    '�߂�l��Null�ł��邽�߃��b�Z�[�W�\��
    If IsNull(filePathArr) Then MsgBox "�t�@�C���p�X�z��̎擾�Ɏ��s���܂����i�����͑��s���܂��j�B"
    
    'funcSample01�̌Ăяo���@���݂���V�[�g���̈����ŌĂяo��
    filePathArr = funcSample01("FilePath")
    '�߂�l��Null�ł͂Ȃ����ߎ��s�̕\���Ȃ�
    If IsNull(filePathArr) Then MsgBox "�t�@�C���p�X�z��̎擾�Ɏ��s���܂����i�����͑��s���܂��j�B"
    
    '���ꂼ��̃G�N�Z���t�@�C���ɂ��āAfuncSample02���Ăяo��
    For Each filePath In filePathArr
        'funcSample02�̌Ăяo��
        '���ł�A1�Z�����������܂�Ă����ꍇ�̓C�~�f�B�G�C�g�E�B���h�E�Ɏ��s�����t�@�C���p�X���o��
        If Not funcSample02(ThisWorkbook.Path & filePath) Then
            Debug.Print "�������ݎ��s�t�@�C���F" & filePath
        End If
    Next filePath
    
    
    '������funcSample01,funcSample02�ȂǂŃL���b�`�ł��Ȃ������z��O�̃G���[��
    '�@�@�@���̃v���V�[�W����ErrorHandler�ŃL���b�`����܂��B
    
ExitHandler:
    
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrorHandler:

    MsgBox "�V�X�e���G���[���������܂����B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "�}�N��"
        
    GoTo ExitHandler
        
End Sub




'******************************************************************************************
'*�֐���    �F��O����Function�v���V�[�W������(1)
'*�@�\      �F�t�@�C���p�X�̕������z��Ƃ��Ď擾
'*����      �F���̃t�@�C���̃V�[�g�̖��O
'*�߂�l    �F������̔z�� > ����I���ANull > �ُ�I��
'******************************************************************************************
Public Function funcSample01(ByVal wsName As String) As Variant
    
    '�萔
    Const FUNC_NAME As String = "funcSample01"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler

    funcSample01 = Null
    
    '�w�肳�ꂽ�V�[�g��A1�Z������A3�Z���܂ł̒l��z��Ƃ��Ď擾����
    With ThisWorkbook.Worksheets(wsName)
        funcSample01 = .Range("A1:A3").Value
    End With

ExitHandler:
    
    
    Exit Function
    
ErrorHandler:

    MsgBox "�V�X�e���G���[���������܂����B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "�}�N��"
        
    GoTo ExitHandler
        
End Function



'******************************************************************************************
'*�֐���    �F��O����Function�v���V�[�W������(2)
'*�@�\      �F�w�肳�ꂽ�p�X�̃G�N�Z���t�@�C�����J��
'               �ꖇ�ڂ̃V�[�g��A1�Z���Ɏ�������������
'               �񖇖ڂ̃V�[�g�����݂���΁A�񖇖ڂ�A1�Z���Ɂu�����v�Ə�������
'*����      �F�G�N�Z���t�@�C���̃p�X
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
Public Function funcSample02(ByVal filePath As String) As Boolean
    
    '�萔
    Const FUNC_NAME As String = "funcSample02"
    
    '�ϐ�
    Dim wb As Workbook
    
    On Error GoTo ErrorHandler

    funcSample02 = False
    
    Set wb = Workbooks.Open(filePath)
    
    
    With wb
        '�ꖇ�ڂ̃V�[�g��A1�Z���Ɏ�������������
        '���ł�A1�Z���ɕ������������܂�Ă����ꍇ�̓G���[�ƂȂ�i�ُ�I���j
        If Trim(.Worksheets(1).Range("A1").Value) <> "" Then Err.Raise 1000, , "A1�Z���ɂ��łɒl�����݂��܂��B"
        .Worksheets(1).Range("A1").Value = Now
        
        '�񖇖ڂ̃V�[�g�����݂��Ȃ���ΏI���i����I���j
        If .Worksheets.Count < 2 Then GoTo TruePoint
        
        '�񖇖ڂ�A1�Z���Ɂu�����v�Ə�������
        .Worksheets(2).Range("A1").Value = "����"
        
    End With
    

TruePoint:
    
    '�V�[�g�̕ۑ�
    wb.Save
    
    funcSample02 = True

ExitHandler:
    
    '����I�����ł��G���[���N�����ꍇ�ł��A�K���u�b�N�����
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    
    Exit Function
    
ErrorHandler:

    MsgBox "�V�X�e���G���[���������܂����B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "�}�N��"
        
    GoTo ExitHandler
        
End Function



