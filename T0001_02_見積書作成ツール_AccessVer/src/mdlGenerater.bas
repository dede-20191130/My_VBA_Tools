Attribute VB_Name = "mdlGenerater"
'@Folder("Database")
Option Compare Database
Option Explicit

'**************************
'*�N���X��������
'**************************

'�萔��
Private Const SOURCE_NAME As String = ""

Public Enum eTypeSettingForm
    consumptionTax = 1
    numUnit = 2
    company = 3
    member = 4
    shoItem = 5
End Enum

Public Enum eTypeOperate
    opeRegister = 16
    opeEdit = 32
End Enum

Public Enum eTypeRegiOrEditForm
    consumptionTaxRegi = 17
    consumptionTaxEdit = 33
    numUnitRegi = 18
    numUnitEdit = 34
    companyRegi = 19
    companyEdit = 35
    memberRegi = 20
    memberEdit = 36
    shoItemRegi = 21
    shoItemEdit = 37
End Enum

Public Enum eTypeRcdSlctForm
    company = 1
    member = 2
End Enum

'�ϐ���



'******************************************************************************************
'*getter/setter��
'******************************************************************************************





'******************************************************************************************
'*�֐���    �FgeneObjSettingForm
'*�@�\      �F�Z�b�e�B���OForm�����I�u�W�F�N�g����
'*����      �F���
'*�߂�l    �F�I�u�W�F�N�g
'******************************************************************************************
Public Function geneObjSettingForm(ByVal argType As eTypeSettingForm) As clsAbsSettingForm
    
    '�萔
    Const FUNC_NAME As String = "geneObjSettingForm"
    
    '�ϐ�
    
    Select Case argType
    Case eTypeSettingForm.consumptionTax
        Set geneObjSettingForm = New clsSettingFormCnsmTx
    Case eTypeSettingForm.numUnit
        Set geneObjSettingForm = New clsSettingFormNmUnt
    Case eTypeSettingForm.company
        Set geneObjSettingForm = New clsSettingFormCompany
    Case eTypeSettingForm.member
        Set geneObjSettingForm = New clsSettingFormMember
    Case eTypeSettingForm.shoItem
        Set geneObjSettingForm = New clsSettingFormShoItem
    Case Else
        err.Raise eNumCustomErr.wrongArgs, , MSG_ERR_WRONG_ARGS
    End Select
        
    
ExitHandler:

    Exit Function
    
        
End Function





'******************************************************************************************
'*�@�\      �F�Z�b�e�B���OForm�����I�u�W�F�N�g����
'*����      �F���
'*�߂�l    �F�I�u�W�F�N�g
'******************************************************************************************
Public Function geneObjRegiOrEdirForm(ByVal argType As eTypeRegiOrEditForm) As clsAbsRegiOrEditForm
    
    '�萔
    Const FUNC_NAME As String = "geneObjRegiOrEdirForm"
    
    '�ϐ�
    
    Select Case argType
    Case eTypeRegiOrEditForm.consumptionTaxRegi
        Set geneObjRegiOrEdirForm = New clsRegiOrEdirFormCnsmTxRegi
    Case eTypeRegiOrEditForm.consumptionTaxEdit
        Set geneObjRegiOrEdirForm = New clsRegiOrEdirFormCnsmTxEdit
    Case eTypeRegiOrEditForm.numUnitRegi
        Set geneObjRegiOrEdirForm = New clsRegiOrEdirFormNmUntRegi
    Case eTypeRegiOrEditForm.numUnitEdit
        Set geneObjRegiOrEdirForm = New clsRegiOrEdirFormNmUntEdit
    Case eTypeRegiOrEditForm.companyRegi
        Set geneObjRegiOrEdirForm = New clsRegiOrEdirFormCompanyRegi
    Case eTypeRegiOrEditForm.companyEdit
        Set geneObjRegiOrEdirForm = New clsRegiOrEdirFormCompanyEdit
    Case eTypeRegiOrEditForm.memberRegi
        Set geneObjRegiOrEdirForm = New clsRegiOrEdirFormMemberRegi
    Case eTypeRegiOrEditForm.memberEdit
        Set geneObjRegiOrEdirForm = New clsRegiOrEdirFormMemberEdit
    Case eTypeRegiOrEditForm.shoItemRegi
        Set geneObjRegiOrEdirForm = New clsRegiOrEdirFormShoItemRegi
    Case eTypeRegiOrEditForm.shoItemEdit
        Set geneObjRegiOrEdirForm = New clsRegiOrEdirFormShoItemEdit
    Case Else
        err.Raise eNumCustomErr.wrongArgs, , MSG_ERR_WRONG_ARGS
    End Select
        
    
ExitHandler:

    Exit Function
    
        
End Function





'******************************************************************************************
'*�֐���    �FgeneObjSettingForm
'*�@�\      �F�Z�b�e�B���OForm�����I�u�W�F�N�g����
'*����      �F���
'*�߂�l    �F�I�u�W�F�N�g
'******************************************************************************************
Public Function geneObjRcdSlctForm(ByVal argType As eTypeRcdSlctForm) As clsAbsRcdSlctForm
    
    '�萔
    Const FUNC_NAME As String = "geneObjRcdSlctForm"
    
    '�ϐ�
    
    Select Case argType
    Case eTypeRcdSlctForm.company
        Set geneObjRcdSlctForm = New clsRcdSlctFormCompany
    Case eTypeRcdSlctForm.member
        Set geneObjRcdSlctForm = New clsRcdSlctFormMember
    Case Else
        err.Raise eNumCustomErr.wrongArgs, , MSG_ERR_WRONG_ARGS
    End Select
        
    
ExitHandler:

    Exit Function
    
        
End Function
