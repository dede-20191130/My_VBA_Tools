Attribute VB_Name = "mdlFile"
'@Folder("Database")
Option Compare Database
Option Explicit

'**************************
'*�t�@�C�����샂�W���[��
'**************************

'�萔��
Private Const SOURCE_NAME As String = "mdlFile"

'�ϐ���



'******************************************************************************************
'*getter/setter��
'******************************************************************************************




'******************************************************************************************
'*�@�\      �F���Ϗ��e���v���[�g���w�肵���t�@�C���p�X�Ƃ��ĕۑ�
'*����      �FDatabase
'*����      �F�t�@�C���p�X
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
Public Function saveBookEstmTmpl(ByVal daoDB As Database, ByVal filePath As String) As Boolean
    
    '�萔
    Const FUNC_NAME As String = "saveBookEstmTmpl"
    
    '�ϐ�
    Dim wrs As New clsWrappedRecordSet
    
    On Error GoTo ErrorHandler

    saveBookEstmTmpl = False
    
    '�e���v���[�g���擾
    Set wrs.varRecordset = daoDB.OpenRecordset( _
        "SELECT FILE FROM " & _
        TBL_M_FILE & myVBVacant & _
        "WHERE FILENAME = 'template';" _
    )
       
    '�ۑ�
    With wrs.varRecordset
        .MoveFirst
        .Fields("FILE").VALUE.Fields("FileData").SaveToFile filePath
    End With
    
    
TruePoint:

    saveBookEstmTmpl = True

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

