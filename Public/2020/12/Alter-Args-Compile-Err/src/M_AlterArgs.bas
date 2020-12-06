Attribute VB_Name = "M_AlterArgs"
Option Explicit


'******************************************************************************************
'*�֐���    �F�Ăяo�����֐�No.1
'*�@�\      �F
'*����      �F
'******************************************************************************************
Public Sub callingFunc01()
    
    '�萔
    Const FUNC_NAME As String = "callingFunc01"
    
    '�ϐ�
    Dim result As Long
    
    On Error GoTo ErrorHandler

    If Not calledFunc(result, 13.54, 28.3) Then GoTo ExitHandler
    result = result + 1000
    Debug.Print result 'output 1383
    
ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "�}�N��"
        
    GoTo ExitHandler
        
End Sub




'******************************************************************************************
'*�֐���    �F�Ăяo�����֐�No.2
'*�@�\      �F
'*����      �F
'******************************************************************************************
Public Sub callingFunc02()
    
    '�萔
    Const FUNC_NAME As String = "callingFunc02"
    
    '�ϐ�
    Dim result As Long
    
    On Error GoTo ErrorHandler

    If Not calledFunc(result, 12.5, 33.33) Then GoTo ExitHandler
    result = result + 5000
    Debug.Print result 'output 5416
    
ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "�}�N��"
        
    GoTo ExitHandler
        
End Sub





'******************************************************************************************
'*�֐���    �F�Ăяo�����֐�
'*�@�\      �F�����̂Q���̏�Z���s���A�����_��؂�̂Ă������𓾂�B
'*�@�@�@�@�@�@���������ɐ��������Z���A�Ԃ�
'*����      �F�Q�Ɠn���Ō��ʂ�ԋp����ϐ�
'*����      �F��Z����鐔�l�P
'*����      �F��Z����鐔�l�Q
'*����      �F���Z����鐮���l
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
#If DEVELOP_MODE Then
    Public Function calledFunc(ByRef returnNum As Long, ByVal num01 As Double, ByVal num02 As Double, Optional ByVal addNum As Long = 0) As Boolean
#Else
    Public Function calledFunc(ByRef returnNum As Long, ByVal num01 As Double, ByVal num02 As Double, ByVal addNum As Long) As Boolean
#End If

    
    '�萔
    Const FUNC_NAME As String = "calledFunc"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler

    calledFunc = False
    
    returnNum = Int(num01 * num02)
    returnNum = returnNum + addNum


TruePoint:

    calledFunc = True

ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "�}�N��"
        
    GoTo ExitHandler
        
End Function


