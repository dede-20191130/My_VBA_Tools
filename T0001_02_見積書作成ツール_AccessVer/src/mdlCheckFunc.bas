Attribute VB_Name = "mdlCheckFunc"
'@Folder("Database")
Option Compare Database
Option Explicit

'**************************************
'*�`�F�b�N�֐����W���[��
'**************************************


'�萔



'�ϐ�


'******************************************************************************************
'*�֐���    �FcheckWhetherControlsVacant
'*�@�\      �F�����̃R���g���[���̒l�̂����A�󗓂ł�����̂����݂��邩�ǂ����𔻒肷��
'*����(1)   �F�󗓔��茋��
'*����(2)   �F�Ώۂ̃R���g���[���̒l�@������
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
Public Function checkWhetherControlsVacant( _
       ByRef isExists As Boolean, _
       ParamArray pCtlVals() As Variant _
       ) As Boolean
    
    '�萔
    Const FUNC_NAME As String = "checkWhetherControlsVacant"
    
    '�ϐ�
    Dim ctlVal As Variant
    
    On Error GoTo ErrorHandler

    checkWhetherControlsVacant = False
    isExists = False
    
    For Each ctlVal In pCtlVals
        If Trim(Nz(ctlVal, "")) = "" Then isExists = True: Exit For
    Next ctlVal

    checkWhetherControlsVacant = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & err.Number & vbNewLine & _
           err.description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Function


'******************************************************************************************
'*�֐���    �FcheckType
'*�@�\      �F�^�`�F�b�N
'*����      �F�]���Ώ�
'*����      �F�^
'*����      �F���ʕԋp�p
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
Public Function checkType( _
    ByVal tgtVal As Variant, _
    ByVal pDataTypeEnum As DataTypeEnum, _
    ByRef isErrorOfType As Boolean _
) As Boolean
    
    '�萔
    Const FUNC_NAME As String = "checkType"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler

    checkType = False
    isErrorOfType = False
    
    '�^�`�F�b�N�֐��Ăяo��
    Select Case pDataTypeEnum
    Case DataTypeEnum.dbText
        isErrorOfType = Not checkTypeText(tgtVal)
    Case DataTypeEnum.dbInteger, DataTypeEnum.dbLong
        isErrorOfType = Not checkTypeIntegral(tgtVal)
    Case DataTypeEnum.dbSingle, DataTypeEnum.dbDouble
        isErrorOfType = Not checkTypeNum(tgtVal)
    Case DataTypeEnum.dbDate
        isErrorOfType = Not checkTypeDate(tgtVal)
    Case DataTypeEnum.dbCurrency
        isErrorOfType = Not checkTypeCur(tgtVal)
    End Select


    checkType = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & err.Number & vbNewLine & _
           err.description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Function



'******************************************************************************************
'*�֐���    �FcheckTypeText
'*�@�\      �F�^�`�F�b�N �e�L�X�g�^�i�ő�t�B�[���h�T�C�Y255�j
'*����      �F�]���Ώ�
'*�߂�l    �FTrue > �w�肳�ꂽ�^�AFalse > �w�肳�ꂽ�^�ł͂Ȃ�
'******************************************************************************************
Public Function checkTypeText(ByVal tgtVal As Variant) As Boolean
    
    '�萔
    Const FUNC_NAME As String = "checkTypeText"
    
    '�ϐ�
    Dim s As String
    
    On Error GoTo ErrorHandler
    
    checkTypeText = True
    
    s = tgtVal
    If Len(s) > 255 Then checkTypeText = False
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    checkTypeText = False
        
    GoTo ExitHandler
        
End Function



'******************************************************************************************
'*�֐���    �FcheckTypeNum
'*�@�\      �F�^�`�F�b�N �����^�@Integer,Long
'*����      �F�]���Ώ�
'*�߂�l    �FTrue > �w�肳�ꂽ�^�AFalse > �w�肳�ꂽ�^�ł͂Ȃ�
'******************************************************************************************
Public Function checkTypeIntegral(ByVal tgtVal As Variant) As Boolean
    
    '�萔
    Const FUNC_NAME As String = "checkTypeIntegral"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler
    
    checkTypeIntegral = True
    
    If Not IsNumeric(tgtVal) Then checkTypeIntegral = False: GoTo ExitHandler
    If CLng(tgtVal) <> tgtVal Then checkTypeIntegral = False: GoTo ExitHandler
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    checkTypeIntegral = False
        
    GoTo ExitHandler
        
End Function







'******************************************************************************************
'*�֐���    �FcheckTypeNum
'*�@�\      �F�^�`�F�b�N ���l�^
'*����      �F�]���Ώ�
'*�߂�l    �FTrue > �w�肳�ꂽ�^�AFalse > �w�肳�ꂽ�^�ł͂Ȃ�
'******************************************************************************************
Public Function checkTypeNum(ByVal tgtVal As Variant) As Boolean
    
    '�萔
    Const FUNC_NAME As String = "checkTypeNum"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler
    
    checkTypeNum = True
    
    If Not IsNumeric(tgtVal) Then checkTypeNum = False
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    checkTypeNum = False
        
    GoTo ExitHandler
        
End Function




'******************************************************************************************
'*�֐���    �FcheckTypeDate
'*�@�\      �F�^�`�F�b�N ���t�^
'*����      �F�]���Ώ�
'*�߂�l    �FTrue > �w�肳�ꂽ�^�AFalse > �w�肳�ꂽ�^�ł͂Ȃ�
'******************************************************************************************
Public Function checkTypeDate(ByVal tgtVal As Variant) As Boolean
    
    '�萔
    Const FUNC_NAME As String = "checkTypeDate"
    
    '�ϐ�
    Dim d As Date
    
    On Error GoTo ErrorHandler
    
    checkTypeDate = True
    
    d = CDate(tgtVal)
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    checkTypeDate = False
        
    GoTo ExitHandler
        
End Function




'******************************************************************************************
'*�֐���    �FcheckTypeCur
'*�@�\      �F�^�`�F�b�N �ʉ݌^
'*����      �F�]���Ώ�
'*�߂�l    �FTrue > �w�肳�ꂽ�^�AFalse > �w�肳�ꂽ�^�ł͂Ȃ�
'******************************************************************************************
Public Function checkTypeCur(ByVal tgtVal As Variant) As Boolean
    
    '�萔
    Const FUNC_NAME As String = "checkTypeCur"
    
    '�ϐ�
    Dim cur As Currency
    
    On Error GoTo ErrorHandler
    
    checkTypeCur = True
    
    cur = CCur(tgtVal)
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    checkTypeCur = False
        
    GoTo ExitHandler
        
End Function




'******************************************************************************************
'*�@�\      �F�d�b�ԍ��`�F�b�N
'*����      �F�Ώە�����
'*����      �F����
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
Public Function checkTelNum(ByVal tgtVal As Variant, ByRef boolError As Boolean) As Boolean
    
    '�萔
    Const FUNC_NAME As String = "checkTelNum"
    
    '�ϐ�
    Dim objReg As New clsWrappedRegExp
    
    On Error GoTo ErrorHandler

    checkTelNum = False
    
    '�����ƃn�C�t���̂݋��e
    boolError = Not objReg.test(CStr(tgtVal), "^[\d-]+$")

TruePoint:

    checkTelNum = True

ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & err.Number & vbNewLine & _
           err.description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Function

