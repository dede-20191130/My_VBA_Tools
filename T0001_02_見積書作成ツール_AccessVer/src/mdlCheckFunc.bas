Attribute VB_Name = "mdlCheckFunc"
'@Folder("Database")
Option Compare Database
Option Explicit

'**************************************
'*チェック関数モジュール
'**************************************


'定数



'変数


'******************************************************************************************
'*関数名    ：checkWhetherControlsVacant
'*機能      ：複数のコントロールの値のうち、空欄であるものが存在するかどうかを判定する
'*引数(1)   ：空欄判定結果
'*引数(2)   ：対象のコントロールの値　複数可
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function checkWhetherControlsVacant( _
       ByRef isExists As Boolean, _
       ParamArray pCtlVals() As Variant _
       ) As Boolean
    
    '定数
    Const FUNC_NAME As String = "checkWhetherControlsVacant"
    
    '変数
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

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & err.Number & vbNewLine & _
           err.description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Function


'******************************************************************************************
'*関数名    ：checkType
'*機能      ：型チェック
'*引数      ：評価対象
'*引数      ：型
'*引数      ：結果返却用
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function checkType( _
    ByVal tgtVal As Variant, _
    ByVal pDataTypeEnum As DataTypeEnum, _
    ByRef isErrorOfType As Boolean _
) As Boolean
    
    '定数
    Const FUNC_NAME As String = "checkType"
    
    '変数
    
    On Error GoTo ErrorHandler

    checkType = False
    isErrorOfType = False
    
    '型チェック関数呼び出し
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

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & err.Number & vbNewLine & _
           err.description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Function



'******************************************************************************************
'*関数名    ：checkTypeText
'*機能      ：型チェック テキスト型（最大フィールドサイズ255）
'*引数      ：評価対象
'*戻り値    ：True > 指定された型、False > 指定された型ではない
'******************************************************************************************
Public Function checkTypeText(ByVal tgtVal As Variant) As Boolean
    
    '定数
    Const FUNC_NAME As String = "checkTypeText"
    
    '変数
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
'*関数名    ：checkTypeNum
'*機能      ：型チェック 整数型　Integer,Long
'*引数      ：評価対象
'*戻り値    ：True > 指定された型、False > 指定された型ではない
'******************************************************************************************
Public Function checkTypeIntegral(ByVal tgtVal As Variant) As Boolean
    
    '定数
    Const FUNC_NAME As String = "checkTypeIntegral"
    
    '変数
    
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
'*関数名    ：checkTypeNum
'*機能      ：型チェック 数値型
'*引数      ：評価対象
'*戻り値    ：True > 指定された型、False > 指定された型ではない
'******************************************************************************************
Public Function checkTypeNum(ByVal tgtVal As Variant) As Boolean
    
    '定数
    Const FUNC_NAME As String = "checkTypeNum"
    
    '変数
    
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
'*関数名    ：checkTypeDate
'*機能      ：型チェック 日付型
'*引数      ：評価対象
'*戻り値    ：True > 指定された型、False > 指定された型ではない
'******************************************************************************************
Public Function checkTypeDate(ByVal tgtVal As Variant) As Boolean
    
    '定数
    Const FUNC_NAME As String = "checkTypeDate"
    
    '変数
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
'*関数名    ：checkTypeCur
'*機能      ：型チェック 通貨型
'*引数      ：評価対象
'*戻り値    ：True > 指定された型、False > 指定された型ではない
'******************************************************************************************
Public Function checkTypeCur(ByVal tgtVal As Variant) As Boolean
    
    '定数
    Const FUNC_NAME As String = "checkTypeCur"
    
    '変数
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
'*機能      ：電話番号チェック
'*引数      ：対象文字列
'*引数      ：結果
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function checkTelNum(ByVal tgtVal As Variant, ByRef boolError As Boolean) As Boolean
    
    '定数
    Const FUNC_NAME As String = "checkTelNum"
    
    '変数
    Dim objReg As New clsWrappedRegExp
    
    On Error GoTo ErrorHandler

    checkTelNum = False
    
    '数字とハイフンのみ許容
    boolError = Not objReg.test(CStr(tgtVal), "^[\d-]+$")

TruePoint:

    checkTelNum = True

ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & err.Number & vbNewLine & _
           err.description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Function

