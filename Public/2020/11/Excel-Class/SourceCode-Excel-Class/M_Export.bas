Attribute VB_Name = "M_Export"
'@Folder("Database")
Option Compare Database
Option Explicit

'**************************
'*�f�[�^�o�̓��W���[��
'**************************

'�萔



'�ϐ�
Public db As DAO.Database


'******************************************************************************************
'*getter/setter
'******************************************************************************************


'******************************************************************************************
'*�֐���    �FgetTableHeader
'*�@�\      �F�e�[�u���̃w�b�_�[������z����擾
'*����      �F�e�[�u����
'*����      �F���ʕԋp�p
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
Public Function getTableHeader(ByVal tblName As String, ByRef pArrHeader() As String) As Boolean
    
    '�萔
    Const FUNC_NAME As String = "getTableHeader"
    
    '�ϐ�
    Dim i As Long
    
    On Error GoTo ErrorHandler

    getTableHeader = False
    Erase pArrHeader
    
    With db.TableDefs(tblName)
        ReDim pArrHeader(0 To .Fields.Count - 1)
        For i = 0 To .Fields.Count - 1
            pArrHeader(i) = .Fields(i).Name
        Next
    End With
    
    getTableHeader = True
    
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


'******************************************************************************************
'*�֐���    �FgetTableDataBySQL
'*�@�\      �FSQL�������蓾��ꂽ���R�[�h�Z�b�g�̃f�[�^��񎟌��z��Ƃ��Ď擾
'*����      �F�Ώ�SQL������
'*����      �F���ʕԋp�p
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
Public Function getTableDataBySQL(ByVal sql As String, ByRef arrData() As Variant) As Boolean
    
    '�萔
    Const FUNC_NAME As String = "getTableDataBySQL"
    
    '�ϐ�
    Dim rs As DAO.Recordset
    Dim i As Long
    Dim j As Long
    
    On Error GoTo ErrorHandler

    getTableDataBySQL = False
    Erase arrData
    
    Set rs = db.OpenRecordset(sql)
    With rs
        If .EOF Then GoTo TruePoint
        .MoveLast
        ReDim arrData(0 To .RecordCount - 1, 0 To .Fields.Count - 1)
        .MoveFirst
        
        i = 0
        Do Until .EOF
            For j = 0 To .Fields.Count - 1
                arrData(i, j) = .Fields(j).Value
            Next j
            i = i + 1
            .MoveNext
        Loop
    End With

TruePoint:

    getTableDataBySQL = True
    
ExitHandler:
    
    If Not rs Is Nothing Then rs.Clone: Set rs = Nothing
    
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




'******************************************************************************************
'*�֐���    �FpostDataToSheet
'*�@�\      �F�V�[�g�ɔz��f�[�^��]�L����
'*����      �F�ΏۃV�[�g
'*����      �F�V�[�g��
'*����      �F�w�b�_�[�z��
'*����      �F�f�[�^�z��
'*�߂�l    �FTrue > ����I���AFalse > �ُ�I��
'******************************************************************************************
Public Function postDataToSheet( _
    ByVal tgtSheet As Object, _
    ByVal sheetName As String, _
    ByVal pArrHeader As Variant, _
    ByVal pArrData As Variant _
) As Boolean
    
    '�萔
    Const FUNC_NAME As String = "postDataToSheet"
    
    '�ϐ�
    
    On Error GoTo ErrorHandler

    postDataToSheet = False
    
    With tgtSheet
        .Name = sheetName
        .Range(.cells(1, 1), .cells(1, UBound(pArrHeader) - LBound(pArrHeader) + 1)).Value = pArrHeader
        .Range(.cells(2, 1), .cells(UBound(pArrData, 1) - LBound(pArrData, 1) + 2, UBound(pArrData, 2) - LBound(pArrData, 2) + 1)).Value = pArrData
        '�r��
        .Range(.cells(1, 1), .cells(UBound(pArrData, 1) - LBound(pArrData, 1) + 2, UBound(pArrData, 2) - LBound(pArrData, 2) + 1)).Borders.LineStyle = xlContinuous
        '�񕝒���
        .Range(.Columns(1), .Columns(UBound(pArrHeader) - LBound(pArrHeader) + 1)).AutoFit
    End With

    postDataToSheet = True
    
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



'******************************************************************************************
'*�֐���    �FgetJsonFromAPI
'*�@�\      �F�w��URL��API����Json��������擾
'*����      �FURL
'*�߂�l    �FJson������iParse�O�j
'******************************************************************************************
Public Function getJsonFromAPI(URL As String) As String

    '�萔
    Const FUNC_NAME As String = "getJsonFromAPI"
    
    '�ϐ�
    Dim objXMLHttp As Object
    
    On Error GoTo ErrorHandler

    getJsonFromAPI = ""
    
    Set objXMLHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        objXMLHttp.Open "GET", URL, False
        objXMLHttp.Send


    getJsonFromAPI = objXMLHttp.responseText
    
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

