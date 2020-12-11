VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableCreater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Class")
Option Explicit

'**************************
'*TableCreater
'**************************


'�萔
Private Const HEADER_COLOR = 15917529            '�w�b�_�[�Z���̐F

'�ϐ�
Private myRange As Range                         '�͈�
Private myColumnSubTotal As Long                 '���v��ԍ�


'******************************************************************************************
'*getter/setter
'******************************************************************************************


Public Property Set Range(ByVal pRng As Range)
    Set myRange = pRng
End Property


Public Property Let ColumnSubTotal(ByVal num As Long)
    '�͈͕ϐ����ݒ肳��Ă��Ȃ���ԂŌďo�͋֎~
    If myRange Is Nothing Then Err.Raise 1000, , "�͈͂��ݒ肳��Ă��܂���B"
    '�������͈̓I�u�W�F�N�g�̗�ԍ����O�̗�̔ԍ��ł���΃G���[
    If num < myRange.EntireColumn(1).Column Or myRange.EntireColumn(myRange.EntireColumn.Count).Column < num Then Err.Raise 1001, , "��ԍ��̎w�肪�s���ł��B"
    '�l��ݒ�
    myColumnSubTotal = num
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
    
    '�{�N���X�ł͓��ɂ�邱�Ƃ͂Ȃ��B

ExitHandler:

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
'*�֐���    �FdrawLines
'*�@�\      �F�r��������
'*����      �F
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
Public Function drawLines() As Boolean
    
    '�萔
    Const FUNC_NAME As String = "drawLines"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler

    drawLines = False
    
    '�͈͕ϐ����ݒ肳��Ă��Ȃ���ԂŌďo�͋֎~
    If myRange Is Nothing Then Err.Raise 1000, , "�͈͂��ݒ肳��Ă��܂���B"
    
    '�r��������
    myRange.Borders.LineStyle = xlContinuous

    drawLines = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "TableCreater"
        
    GoTo ExitHandler
        
End Function


'******************************************************************************************
'*�֐���    �FsetStyleForHeader
'*�@�\      �F�w�b�_�[�̋����̂��߂̃X�^�C���ύX���s��
'               �w�b�_�[�͗^����ꂽ�͈͂̈�s�ڂ̃Z���Ƃ���
'*����      �F
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
Public Function setStyleForHeader() As Boolean
    
    '�萔
    Const FUNC_NAME As String = "setStyleForHeader"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler

    setStyleForHeader = False
    
    '�͈͕ϐ����ݒ肳��Ă��Ȃ���ԂŌďo�͋֎~
    If myRange Is Nothing Then Err.Raise 1000, , "�͈͂��ݒ肳��Ă��܂���B"
    
    '�X�^�C���ύX
    With myRange.Rows(1)
        '�͈͂̈�s�ڂ̃Z���̐F��ύX
        .Interior.color = HEADER_COLOR
        '�͈͂̈�s�ڂ𑾎��ɂ���
        .Font.Bold = True
        '�͈͂̈�s�ڂ𒆉������ɂ���
        .HorizontalAlignment = xlCenter
    End With
        
    
    setStyleForHeader = True
    
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "TableCreater"
        
    GoTo ExitHandler
        
End Function




'******************************************************************************************
'*�֐���    �FcalcTotalFromSubTotal
'*�@�\      �F���v�񂩂獇�v���v�Z���o�͂���
'*����      �F
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
Public Function calcTotalFromSubTotal() As Boolean
    
    '�萔
    Const FUNC_NAME As String = "calcTotalFromSubTotal"
    
    '�ϐ�
    Dim sumVal As Long
    Dim cell As Range
    Dim subTotalOrder As Long
    
    On Error GoTo ErrorHandler

    calcTotalFromSubTotal = False
    
    '�͈͕ϐ����ݒ肳��Ă��Ȃ���ԂŌďo�͋֎~
    If myRange Is Nothing Then Err.Raise 1000, , "�͈͂��ݒ肳��Ă��܂���B"
    '���v�񂪐ݒ肳��Ă��Ȃ���ԂŌďo�͋֎~
    If myColumnSubTotal = 0 Then Err.Raise 1002, , "���v�񂪐ݒ肳��Ă��܂���B"
        
    '���v�񂪔͈̗͂�̉��Ԗڂł��邩���Z�o
    subTotalOrder = myColumnSubTotal - myRange(1).Column + 1
        
    '���v�l�v�Z�i�w�b�_�[��͊܂܂Ȃ��j
    For Each cell In myRange.Columns(subTotalOrder).Cells.Offset(1).Resize(myRange.Columns(subTotalOrder).Cells.Offset(1).Cells.Count - 1)
        '���l�̏ꍇ�̂݉��Z
        If IsNumeric(cell.Value) Then sumVal = sumVal + cell.Value
    Next cell
    If sumVal = 0 Then GoTo TruePoint
    
    '���v�l�����v��̍ŉ��[�Z���̉��̃Z���ɋL��
    With myRange.Columns(subTotalOrder).Rows(myRange.Columns(subTotalOrder).Cells.Count).Offset(1)
        .Value = sumVal
        '���v���x���Z��
        With .Offset(, -1)
            '���x�����L��
            .Value = "���v"
            '���v���x���Z���ƍ��v�Z���Ɍr��������
            .Resize(, .Columns.Count + 1).Borders.LineStyle = xlContinuous
        End With
        
    End With
       
TruePoint:
       
    calcTotalFromSubTotal = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "TableCreater"
        
    GoTo ExitHandler
        
End Function

