Attribute VB_Name = "mdlConst"
'@Folder("Database")
Option Compare Database
Option Explicit

'**************************************
'*Const���W���[��
'**************************************


'�萔

'�ϐ�


'*****�c�[���v���p�e�B
Public Const TOOL_NAME As String = "T0001_02_���Ϗ��쐬�c�[��"

'*****VB�̃J�X�^�}�C�Y�萔
Public Const myVBVacant As String = " "
Public Const myVBUL As String = "_"
Public Const myVBSglQte As String = "'"


'*****ACTIVEX
Public Const SCRIPTING_DICTIONARY As String = "scripting.dictionary"


'*****�e�[�u��
Public Const TBL_M_BSC_TAX As String = "M_��b�f�[�^_����ŗ�"
Public Const TBL_M_BSC_UNIT As String = "M_��b�f�[�^_���ʒP��"
Public Const TBL_M_MEMBER As String = "M_��b�f�[�^_�l���f�[�^"
Public Const TBL_M_PROD As String = "M_��b�f�[�^_���i�f�[�^"
Public Const TBL_M_ORG As String = "M_��b�f�[�^_������Ѓf�[�^"
Public Const TBL_M_FILE As String = "M_�t�@�C��"
Public Const TBL_T_ESTM As String = "T_����_�\�쐬�p�f�[�^"
Public Const TBL_T_ESTM_DTL As String = "T_���Ϗ����ڃf�[�^"
Public Const TBL_W_ESTM As String = "W_����_�\�쐬�p�f�[�^"

'*****�N�G��
Public Const QRY_QSR01 As String = "QRS01_�l��"
Public Const QRY_Q_MAX_W_ESTM As String = "Q01_MAX_W_����_�\�쐬�p�f�[�^"


'*****�t�H�[��
Public Const FormName_01 As String = "F01_Init"
Public Const FormName_02 As String = "F02_���j���["
Public Const FormName_03 As String = "F03_�ݒ�"
Public Const FormName_03_SUB01 As String = "F03_�ݒ�_SUB01_��b�f�[�^_����ŗ�"
Public Const FormName_03_SUB02 As String = "F03_�ݒ�_SUB02_��b�f�[�^_���ʒP��"
Public Const FormName_03_SUB03 As String = "F03_�ݒ�_SUB03_��b�f�[�^_������Ѓf�[�^"
Public Const FormName_03_SUB04 As String = "F03_�ݒ�_SUB04_��b�f�[�^_�l���f�[�^"
Public Const FormName_03_SUB05 As String = "F03_�ݒ�_SUB05_��b�f�[�^_���i�f�[�^"
Public Const FormName_04 As String = "F04_�o�^_�ҏW"
Public Const FormName_04_SUB01 As String = "F04_�o�^_�ҏW_SUB01_��b�f�[�^_����ŗ�"
Public Const FormName_04_SUB02 As String = "F04_�o�^_�ҏW_SUB02_��b�f�[�^_���ʒP��"
Public Const FormName_04_SUB03 As String = "F04_�o�^_�ҏW_SUB03_��b�f�[�^_������Ѓf�[�^"
Public Const FormName_04_SUB04 As String = "F04_�o�^_�ҏW_SUB04_��b�f�[�^_�l���f�[�^"
Public Const FormName_04_SUB05 As String = "F04_�o�^_�ҏW_SUB05_��b�f�[�^_���i�f�[�^"
Public Const FormName_05 As String = "F05_���Ϗ����ڐݒ�"
Public Const FormName_06 As String = "F06_���R�[�h�I��"
Public Const FormName_07 As String = "F07_���R�[�h�I��_���Ϗ�����"
Public Const FormName_08 As String = "F08_���Ϗ�_�\�쐬"
Public Const FormName_09 As String = "F09_���Ϗ�_���ڃ��X�g���R�[�h�ǉ�"
Public Const FormName_10 As String = "F10_�f�[�^�V�[�g�\��"
Public Const FormName_11 As String = "F11_���Ϗ�_���ڃ��X�g���R�[�h�폜"

'*****���O��`�̃L�[
Public Const STR_NAME_RANGE_PRODUCT_TABLE_HEDD As String = "���i�e�[�u��_�w�b�_�["
Public Const STR_NAME_RANGE_TMPL_OTHER_NOTES As String = "�e���v���[�g_���̑����l�i�[�Z��"
Public Const STR_NAME_RANGE_TMPL_COMPANY As String = "�e���v���[�g_��Ж��i�[�Z��"
Public Const STR_NAME_RANGE_TMPL_ESTIMATE_DATE As String = "�e���v���[�g_���ϓ��i�[�Z��"
Public Const STR_NAME_RANGE_TMPL_NUMBER As String = "�e���v���[�g_���ϔԍ��i�[�Z��"
Public Const STR_NAME_RANGE_TMPL_PAYMENT_TERMS As String = "�e���v���[�g_�x�������i�[�Z��"
Public Const STR_NAME_RANGE_TMPL_PERSON_IN_CHARGE As String = "�e���v���[�g_�S���Ҋi�[�Z��"
Public Const STR_NAME_RANGE_TMPL_DELIVERY_DATE As String = "�e���v���[�g_�[���i�[�Z��"
Public Const STR_NAME_RANGE_TMPL_EXPIRATION_DATE As String = "�e���v���[�g_�L�������i�[�Z��"
Public Const STR_NAME_RANGE_TMPL_TAX As String = "�e���v���[�g_����Ŋi�[�Z��"

'*****���b�Z�[�W
Public Const MESSAGE_TITLE_NOTICE As String = "����"
Public Const MESSAGE_TITLE_WARNING As String = "�x��"
Public Const MESSAGE_TITLE_ERROR As String = "�G���["

'�x��
Public Const MESSAGE_EXIST_BLANK As String = "���͗��ɋ󗓂����݂��܂��B"

'*****�G���[�ԍ��E���b�Z�[�W
Public Enum eNumCustomErr
    wrongArgs = 2000
End Enum

Public Const MSG_ERR_WRONG_ARGS As String = "�������s���ł��B"


'*****�Œ蕶��
Public Const REGISTER As String = "�o�^"
Public Const EDIT As String = "�ҏW"
