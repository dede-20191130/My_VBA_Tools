Attribute VB_Name = "mdlDevelop"
'@Folder("Database")
Option Compare Database
Option Explicit

#If DEBUG_MODE Then



Sub debugPrintDev(x As Variant)
    #If Not CBool(DEBUG_MODE) Then
        MsgBox "DEBUG_MODE�ł͂Ȃ����߂��̃R�[�h���폜���Ă�������"
        Exit Sub
    #End If
    Debug.Print Now & " " & CStr(x)
End Sub



'******************************************************************************************
'*�֐���    �FexportCodesSQLs
'*�@�\      �F���W���[���E�N���X�̃R�[�h�y�уN�G����SQL�̏o��
'*����      �F
'******************************************************************************************
Sub exportCodesSQLs()
    
    '�萔
    Const FUNC_NAME As String = "exportCodesSQLs"
    
    '�ϐ�
    Dim outputDir As String
    Dim vbcmp As Object
    Dim fileName As String
    Dim ext As String
    Dim qry As QueryDef
    Dim qName As String
    
    
    
    On Error GoTo ErrorHandler
    
    outputDir = _
        Access.CurrentProject.Path & _
        "\" & _
        "src_" & _
        Left(Access.CurrentProject.Name, InStrRev(Access.CurrentProject.Name, ".") - 1)
    If Dir(outputDir) = "" Then MkDir outputDir
    
    '���W���[���E�N���X�̏o��
    For Each vbcmp In VBE.ActiveVBProject.VBComponents
        With vbcmp
            '�g���q
            Select Case .Type
            Case 1
                ext = ".bas"
            Case 2, 100
                ext = ".cls"
            Case 3
                ext = ".frm"
            End Select
                        
            fileName = .Name & ext
            fileName = gainStrNameSafe(fileName) '�t�@�C�����Ɏg�p�ł��Ȃ�������u��
            If fileName = "" Then GoTo ExitHandler
            
            'output
            .Export outputDir & "\" & fileName
            
        End With
    Next vbcmp
    
    'SQL�̏o��
    With CreateObject("Scripting.FileSystemObject")
        For Each qry In CurrentDb.QueryDefs
            Do
                qName = gainStrNameSafe(qry.Name) '�t�@�C�����Ɏg�p�ł��Ȃ�������u��
                If qName = "" Then GoTo ExitHandler
                
                If qName Like "Msys*" Then Exit Do '�V�X�e���֘A�N�G���͏��O
                
                With .CreateTextFile(outputDir & "\" & qName & ".sql")
                    .write qry.SQL
                    .Close
                End With
            Loop While False
        Next qry
    End With

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & err.Number & vbNewLine & _
           err.description, vbCritical, "�}�N��"
        
    GoTo ExitHandler
        
End Sub




'******************************************************************************************
'*�֐���    �FgainStrNameSafe
'*�@�\      �F�t�@�C�����Ɏg�p�ł��Ȃ��������A���_�[�X�R�A�ɒu������
'*����      �F�Ώۂ̕�����
'*�߂�l    �F�u���㕶����
'******************************************************************************************
Public Function gainStrNameSafe(ByVal s As String) As String
    
    '�萔
    Const FUNC_NAME As String = "gainStrNameSafe"
    
    '�ϐ�
    Dim x As Variant
    
    On Error GoTo ErrorHandler

    gainStrNameSafe = ""
    
    For Each x In Split("\,/,:,*,?,"",<,>,|", ",") '�t�@�C�����Ɏg�p�ł��Ȃ������̔z��
        s = replace(s, x, "_")
    Next x
    
    gainStrNameSafe = s

ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & err.Number & vbNewLine & _
           err.description, vbCritical, "�}�N��"
        
    GoTo ExitHandler
        
End Function




#End If
