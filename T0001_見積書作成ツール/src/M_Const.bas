Attribute VB_Name = "M_Const"
'@Folder("VBAProject")
Option Explicit

'# tool name
Public Const TOOL_NAME As String = "T0001_���Ϗ��쐬�c�[��"

'# �V�[�g���O
Public Const SHEET_NAME_ESTIMATE_DATA As String = "���σf�[�^"
Public Const SHEET_NAME_ESTIMATE_PRODUCT_SET_DATA As String = "���Ϗ��i�Z�b�g�f�[�^"
Public Const SHEET_NAME_PRODUCT_DATA As String = "���i�f�[�^"
Public Const SHEET_NAME_BASIC_DATA As String = "��b�f�[�^"
Public Const SHEET_NAME_TEMPLATE As String = "�e���v���[�g"

'# �V�[�g�I�u�W�F�N�g
Public ws_estimate_data As Worksheet
Public ws_estimate_product_set_data As Worksheet
Public ws_product_data As Worksheet
Public ws_basic_data As Worksheet
Public ws_template As Worksheet

'OLE�I�u�W�F�N�g�Ăяo����
Public Const OP_BUTTON_PDF_EXPORT_OFF As String = "OptionButton_PDF_Export_OFF"
Public Const OP_BUTTON_PDF_EXPORT_ON As String = "OptionButton_PDF_Export_ON"

'# ���O��`�̃L�[
Public Const STR_NAME_RANGE_ESTIMATE_DATA_HEDD As String = "���σf�[�^��_�w�b�_�["
Public Const STR_NAME_RANGE_PRODUCT_DATA_HEDD As String = "���i�f�[�^��_�w�b�_�["
Public Const STR_NAME_RANGE_ESTIMATE_PRODUCT_SET_DATA_HEDD As String = "���Ϗ��i�Z�b�g�f�[�^��_�w�b�_�["
Public Const STR_NAME_RANGE_PRODUCT_TABLE_HEDD As String = "���i�e�[�u��_�w�b�_�["
Public Const STR_NAME_RANGE_ESTIMATE_DATA_SET_NUM_FIELD As String = "���σf�[�^_���Ϗ��i�Z�b�g�ԍ����͗�"
Public Const STR_NAME_RANGE_ESTIMATE_PRODUCT_SET_DATA_CODE_FIELD As String = "���Ϗ��i�Z�b�g�f�[�^_���i�R�[�h���͗�"
Public Const STR_NAME_RANGE_CONSUME_TAX As String = "����ł̒l"
Public Const STR_NAME_RANGE_TMPL_OTHER_NOTES As String = "�e���v���[�g_���̑����l�i�[�Z��"
Public Const STR_NAME_RANGE_TMPL_COMPANY As String = "�e���v���[�g_��Ж��i�[�Z��"
Public Const STR_NAME_RANGE_TMPL_ESTIMATE_DATE As String = "�e���v���[�g_���ϓ��i�[�Z��"
Public Const STR_NAME_RANGE_TMPL_NUMBER As String = "�e���v���[�g_���ϔԍ��i�[�Z��"
Public Const STR_NAME_RANGE_TMPL_PAYMENT_TERMS As String = "�e���v���[�g_�x�������i�[�Z��"
Public Const STR_NAME_RANGE_TMPL_PERSON_IN_CHARGE As String = "�e���v���[�g_�S���Ҋi�[�Z��"
Public Const STR_NAME_RANGE_TMPL_DELIVERY_DATE As String = "�e���v���[�g_�[���i�[�Z��"
Public Const STR_NAME_RANGE_TMPL_EXPIRATION_DATE As String = "�e���v���[�g_�L�������i�[�Z��"
Public Const STR_NAME_RANGE_TMPL_TAX As String = "�e���v���[�g_����Ŋi�[�Z��"

'# �G���[����
Public Const ITEM_KEY_FOR_ERR_MSG As String = "$1"

Public Const ERR_MSG_CREATED_DOCS_MAIN_HEDD As String = "���L�̓��̓~�X�����݂��܂��B" & _
vbLf & _
"�G���[���e�ɏ]���ďC�������Ă��������B"
Public Const ERR_MSG_CREATED_DOCS_MAIN_EACH_HEDD As String = "���L�̃f�[�^�ԍ��̌��Ϗ��͐���ɍ쐬����܂���ł����B"
Public Const ERR_MSG_INVALID_VALUE As String = "�y$1�z" & vbTab & "�L���Ȓl�����͂���Ă��܂���B"
Public Const ERR_MSG_INCONSISTENCY_OF_DATA_NUM As String = "�y$1�z" & vbTab & "�͂��߂̒l�͏I���̒l�ȉ��ɂ��Ă��������B"
Public Const ERR_MSG_ESTIMATE_DATA_MUST_INPUT As String = "�y$1�z" & vbTab & "���σf�[�^�\�̕K�{���ڂ��󗓂ł��B"
Public Const ERR_MSG_ESTIMATE_PRODUCT_SET_DATA_IS_NULL As String = "�y$1�z" & vbTab & "�w�肳�ꂽ���Ϗ��i�Z�b�g�ԍ��ɑΉ�����f�[�^�����݂��܂���B"
Public Const ERR_MSG_INVALID_INDEX As String = "�c�[���̃C���f�b�N�X���ύX����Ă���\��������܂��B" _
& vbLf _
& "�c�[���̃����e�i���X���K�v�ł��B"
Public Const ERR_MSG_CHANGE_EVENT_DUPLICATION As String = "�d������l�͓��͂ł��܂���B"

'# ActiveX�I�u�W�F�N�g�����萔
Public Const STR_ACTIVEX_OBJ_FILE_SYSTEM_OBJ As String = "Scripting.Filesystemobject"
Public Const STR_ACTIVEX_OBJ_DICTIONARY As String = "Scripting.Dictionary"

'�ė��p����N���X�I�u�W�F�N�g
Public obj_set_data As Cls_Set_Data
Public obj_product_data As Cls_Product_Data

'# �J���p�萔
'�C�x���g������
Public Const EVT_DISABLE_FLG As Boolean = False




