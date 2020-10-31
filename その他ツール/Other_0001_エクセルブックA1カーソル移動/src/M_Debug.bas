Attribute VB_Name = "M_Debug"
'@Folder("VBAProject")
Option Explicit


Sub s20201031_2105()
    
    
    With New clsFormatExcel
        Debug.Print Now, _
                    .formatExcel( _
                    "C:\Users\dede2\OneDrive\デスクトップ\tmp\20201031\M_UI.xls", _
                    CreateObject("scripting.filesystemobject") _
                    )
    End With
End Sub



