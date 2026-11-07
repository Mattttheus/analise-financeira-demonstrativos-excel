Attribute VB_Name = "Módulo3"
Sub excluir_registro()
Attribute excluir_registro.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Range("B5:M5").Select
    Selection.ListObject.ListRows(1).Delete
    Range("B5").Select
End Sub
