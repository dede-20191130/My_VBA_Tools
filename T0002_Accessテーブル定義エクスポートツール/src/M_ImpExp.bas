Attribute VB_Name = "M_ImpExp"
'@Folder("Database")
Option Compare Database
Option Explicit

'**************************
'*�C���|�[�g/�G�N�X�|�[�gModule
'**************************

'�萔


'�ϐ�


'******************************************************************************************
'*�֐���    �FexportTableDefTablesMain
'*�@�\      �F�e�[�u����`���e�[�u�����쐬
'*����(1)   �F�Ώ�Access�t�@�C���̃p�X
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
Public Function exportTableDefTablesMain(ByVal dbFilePath As String) As Boolean
    
    '�萔
    Const FUNC_NAME As String = "exportTableDefTablesMain"
    
    '�ϐ�
    Dim selectedDB As DAO.Database
    Dim xlApp As Object
    Dim wb As Object
    Dim tdf As DAO.TableDef
    Dim defArr As Variant
    Dim fstWs As Object
    Dim ws As Object
    
    On Error GoTo ErrorHandler
    
    exportTableDefTablesMain = False
    
    'Access�f�[�^�x�[�X�ɐڑ�
    Set selectedDB = getAccessDB(dbFilePath)
    If selectedDB Is Nothing Then GoTo ExitHandler
    
    '�G�N�Z���u�b�N�J�n
    Set xlApp = CreateObject("Excel.Application")
    With xlApp
        .Visible = False
        .DisplayAlerts = False
        .ScreenUpdating = False
    End With
    Set wb = xlApp.Workbooks.Add
    
    '�����V�[�g
    Set fstWs = wb.Worksheets(1)
    
    '�e�[�u�����ƂɕʃV�[�g�Ƀe�[�u����`���e�[�u�����쐬
    For Each tdf In selectedDB.TableDefs
        Do
            '�V�X�e���e�[�u�����o�͂̕K�v�̂Ȃ��e�[�u���̏ꍇ��continue
            If Left(tdf.Name, 4) = "Msys" Or Left(tdf.Name, 4) = "Usys" Or Left(tdf.Name, 1) = "~" Then Exit Do
            
            '�e�[�u���̒�`���z����擾
            defArr = getTableDefArray(tdf, selectedDB)
            If IsNull(defArr) Then GoTo ExitHandler
            
            '�u�b�N�ŐV�K�V�[�g���쐬
            Set ws = wb.Worksheets.Add
            If Not setWSName(ws, tdf.Name) Then Call Err.Raise(1000, "�V�[�g���w��G���[", "�V�[�g���w��̍ۂɃG���[���������܂����B")
            
            '��`���z����L�����A�񕝒���
            With ws.Range(ws.cells(1, 1), ws.cells(UBound(defArr) - LBound(defArr) + 1, UBound(defArr, 2) - LBound(defArr, 2) + 1))
                .Value = defArr
                .EntireColumn.AutoFit
            End With
            
        Loop While False
    Next tdf
    
    '�����V�[�g�̍폜
    If wb.Worksheets.Count > 1 Then fstWs.Delete
    
    '�u�b�N�ۑ�
    wb.saveas Left( _
              dbFilePath, _
              InStrRev(dbFilePath, ".") - 1 _
              ) & _
                "_�e�[�u����`�ꗗ.xlsx"
    
    '����
    MessageBoxTimeoutA 0&, "�G�N�X�|�[�g����", "�ʒm", vbOKOnly, 0&, 3000
    
    exportTableDefTablesMain = True
    
ExitHandler:
    
    '�N���[�Y
    If Not wb Is Nothing Then wb.Close: Set wb = Nothing
    If Not xlApp Is Nothing Then xlApp.Quit: Set xlApp = Nothing
    If Not selectedDB Is Nothing Then selectedDB.Close: Set selectedDB = Nothing
    
    Set tdf = Nothing
    Set ws = Nothing
    Set fstWs = Nothing
    
    Exit Function
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.description, vbCritical, "�G���["
        
    GoTo ExitHandler
        
End Function


'******************************************************************************************
'*�֐���    �FgetTableDefArray
'*�@�\      �F�e�[�u���̒�`�����擾
'*            ���ځF�t�B�[���h��
'*                  �f�[�^�^
'*                  �T�C�Y
'*                  �K�{���ڂ��ǂ���
'*                  ��L�[�iPK�j
'*                  �O���L�[�iFK�j
'*                  ����
'*
'*����(1)   �F�e�[�u����`
'*�߂�l    �F��`���z��
'******************************************************************************************
Public Function getTableDefArray( _
       ByVal pTdf As DAO.TableDef, _
       ByVal pSelectedDB As DAO.Database _
       ) As Variant
    
    '�萔
    Const FUNC_NAME As String = "getTableDefArray"
    
    '�ϐ�
    Dim defArr() As Variant
    Dim fld As DAO.Field
    Dim i As Long
    Dim dicPKs As Object
    Dim dicFKs As Object
    Dim description As String
    
    On Error GoTo ErrorHandler

    getTableDefArray = Null
    
    '(�e�[�u���̃t�B�[���h�� + 1)�~7�̃T�C�Y�̔z��
    ReDim defArr(0 To pTdf.Fields.Count, 0 To 6)
    
    '�w�b�_�ݒ�
    defArr(0, 0) = "�t�B�[���h��"
    defArr(0, 1) = "�f�[�^�^"
    defArr(0, 2) = "�T�C�Y"
    defArr(0, 3) = "�K�{"
    defArr(0, 4) = "PK"
    defArr(0, 5) = "FK"
    defArr(0, 6) = "����"
    
    '�e�[�u���̂��ׂĂ̎�L�[�ł���t�B�[���h���������Ƃ��Ď擾
    Set dicPKs = getPKs(pTdf)
    If dicPKs Is Nothing Then GoTo ExitHandler
    
    '�e�[�u���̂��ׂĂ̊O���L�[�ł���t�B�[���h���������Ƃ��Ď擾
    Set dicFKs = getFKs(pTdf, pSelectedDB)
    If dicFKs Is Nothing Then GoTo ExitHandler
    
    '�t�B�[���h���ƂɒT��
    For i = 1 To pTdf.Fields.Count
        Set fld = pTdf.Fields(i - 1)
        '�t�B�[���h��
        defArr(i, 0) = fld.Name
        '�f�[�^�^
        defArr(i, 1) = getFieldTypeString(fld.Type)
        '�T�C�Y
        If fld.Type = dbText Then
            defArr(i, 2) = fld.Size
        Else
            defArr(i, 2) = "-"
        End If
        '�K�{���ڂ��ǂ���
        If fld.Required Then defArr(i, 3) = "��"
        '��L�[�iPK�j���ǂ���
        If dicPKs.Exists(fld.Name) Then defArr(i, 4) = "��"
        '�O���L�[�iFK�j���ǂ���
        If dicFKs.Exists(fld.Name) Then defArr(i, 5) = "��"
        '����
        On Error Resume Next
        description = fld.Properties("Description")
        On Error GoTo ErrorHandler
        defArr(i, 6) = description
    Next i


    getTableDefArray = defArr
    
ExitHandler:
    
    Set fld = Nothing
    Set dicFKs = Nothing
    Set dicPKs = Nothing
    
    Exit Function
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.description, vbCritical, "�G���["
        
    GoTo ExitHandler
        
End Function


'******************************************************************************************
'*�֐���    �FgetFieldTypeString
'*�@�\      �F�t�B�[���h�̃f�[�^�^��������擾
'*����(1)   �F�t�B�[���h�^�C�v
'*�߂�l    �F�t�B�[���h�̃f�[�^�^������
'******************************************************************************************
Public Function getFieldTypeString(ByVal pFldTyepNum As Long) As String
    
    '�萔
    Const FUNC_NAME As String = "getFieldTypeString"
    
    '�ϐ�
    Dim strType As String
    
    On Error GoTo ErrorHandler

    strType = ""
    

    Select Case pFldTyepNum
    Case dbBoolean
        strType = "�u�[���^"
    Case dbByte
        strType = "�o�C�g�^"
    Case dbInteger
        strType = "�����^"
    Case dbLong
        strType = "�������^"
    Case dbSingle
        strType = "�P���x���������_�^"
    Case dbDouble
        strType = "�{���x���������_�^"
    Case dbCurrency
        strType = "�ʉ݌^"
    Case dbDate
        strType = "���t/�����^"
    Case dbText
        strType = "�e�L�X�g�^"
    Case dbLongBinary
        strType = "OLE�I�u�W�F�N�g�^"
    Case dbMemo
        strType = "�����^"
    End Select

    getFieldTypeString = strType
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.description, vbCritical, "�G���["
        
    GoTo ExitHandler
        
End Function


'******************************************************************************************
'*�֐���    �FgetPKs
'*�@�\      �F�e�[�u���̎�L�[�ł���t�B�[���h���������Ƃ��Ď擾
'*����(1)   �F�t�B�[���h�^�C�v
'*�߂�l    �F����
'******************************************************************************************
Public Function getPKs(ByVal pTdf As DAO.TableDef) As Object
    
    '�萔
    Const FUNC_NAME As String = "getPKs"
    
    '�ϐ�
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    Dim dic As Object
    
    On Error GoTo ErrorHandler

    Set getPKs = Nothing
    Set dic = CreateObject("Scripting.Dictionary")
    
    
    '�C���f�b�N�X���T��
    For Each idx In pTdf.Indexes
        If idx.Primary = True Then
            For Each fld In idx.Fields
                dic.Add fld.Name, True
            Next
        End If
    Next

    'Return
    Set getPKs = dic
    
ExitHandler:

    Set dic = Nothing

    Exit Function
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.description, vbCritical, "�G���["
        
    GoTo ExitHandler
        
End Function


'******************************************************************************************
'*�֐���    �FgetFKs
'*�@�\      �F�e�[�u���̊O���L�[�ł���t�B�[���h���������Ƃ��Ď擾
'*����(1)   �F
'*�߂�l    �F����
'******************************************************************************************
Public Function getFKs( _
       ByVal pTdf As DAO.TableDef, _
       ByVal pSelectedDB As DAO.Database _
       ) As Object
    
    '�萔
    Const FUNC_NAME As String = "getFKs"
    
    '�ϐ�
    Dim rsRelation As DAO.Recordset
    Dim dic As Object
    
    On Error GoTo ErrorHandler

    Set getFKs = Nothing
    Set dic = CreateObject("Scripting.Dictionary")
    
    '�����[�V�����e�[�u���ɃA�N�Z�X
    Set rsRelation = pSelectedDB.OpenRecordset( _
                     "SELECT szColumn FROM MSysRelationships WHERE szObject =" & _
                     " " & _
                     "'" & _
                     pTdf.Name & _
                     "'" & _
                     ";" _
                     )
    
    With rsRelation
        If .EOF Then Set getFKs = dic: GoTo ExitHandler
        .MoveFirst
        Do Until .EOF
            dic.Add .Fields("szColumn").Value, True
            .MoveNext
        Loop
    End With
    
    'Return
    Set getFKs = dic
    
ExitHandler:
    
    If Not rsRelation Is Nothing Then rsRelation.Close: Set rsRelation = Nothing
        
    Exit Function
    
ErrorHandler:

    MsgBox "�G���[�������������߁A�}�N�����I�����܂��B" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ��F" & Err.Number & vbNewLine & _
           Err.description, vbCritical, "�G���["

        
    GoTo ExitHandler
        
End Function


'******************************************************************************************
'*�֐���    �FsetWSName
'*�@�\      �F�G�N�Z���V�[�g�̖��O���Z�b�g
'*����(1)   �F�G�N�Z���V�[�g
'*����(2)   �F������閼�O
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
Public Function setWSName( _
       ByVal ws As Object, _
       ByVal newName As String _
       ) As Boolean
    
    '�萔
    Const FUNC_NAME As String = "setWSName"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler

    setWSName = False
    
    ws.Name = newName

    setWSName = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    '�V�[�g���Ɏg�p�ł��Ȃ������ł������ꍇ
    ws.Name = "�e�[�u��_" & ws.Parent.Worksheets.Count & "_" & Format(Now, "yyyymmddhhnnss")

    setWSName = True
    GoTo ExitHandler
        
End Function



