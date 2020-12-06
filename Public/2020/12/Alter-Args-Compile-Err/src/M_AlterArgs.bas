Attribute VB_Name = "M_AlterArgs"
Option Explicit


'******************************************************************************************
'*関数名    ：呼び出し元関数No.1
'*機能      ：
'*引数      ：
'******************************************************************************************
Public Sub callingFunc01()
    
    '定数
    Const FUNC_NAME As String = "callingFunc01"
    
    '変数
    Dim result As Long
    
    On Error GoTo ErrorHandler

    If Not calledFunc(result, 13.54, 28.3) Then GoTo ExitHandler
    result = result + 1000
    Debug.Print result 'output 1383
    
ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "マクロ"
        
    GoTo ExitHandler
        
End Sub




'******************************************************************************************
'*関数名    ：呼び出し元関数No.2
'*機能      ：
'*引数      ：
'******************************************************************************************
Public Sub callingFunc02()
    
    '定数
    Const FUNC_NAME As String = "callingFunc02"
    
    '変数
    Dim result As Long
    
    On Error GoTo ErrorHandler

    If Not calledFunc(result, 12.5, 33.33) Then GoTo ExitHandler
    result = result + 5000
    Debug.Print result 'output 5416
    
ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "マクロ"
        
    GoTo ExitHandler
        
End Sub





'******************************************************************************************
'*関数名    ：呼び出される関数
'*機能      ：引数の２数の乗算を行い、小数点を切り捨てし整数を得る。
'*　　　　　　得た整数に整数を加算し、返す
'*引数      ：参照渡しで結果を返却する変数
'*引数      ：乗算される数値１
'*引数      ：乗算される数値２
'*引数      ：加算される整数値
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
#If DEVELOP_MODE Then
    Public Function calledFunc(ByRef returnNum As Long, ByVal num01 As Double, ByVal num02 As Double, Optional ByVal addNum As Long = 0) As Boolean
#Else
    Public Function calledFunc(ByRef returnNum As Long, ByVal num01 As Double, ByVal num02 As Double, ByVal addNum As Long) As Boolean
#End If

    
    '定数
    Const FUNC_NAME As String = "calledFunc"
    
    '変数
    
    On Error GoTo ErrorHandler

    calledFunc = False
    
    returnNum = Int(num01 * num02)
    returnNum = returnNum + addNum


TruePoint:

    calledFunc = True

ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "マクロ"
        
    GoTo ExitHandler
        
End Function


