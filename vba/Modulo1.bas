Attribute VB_Name = "Módulo1"
Sub inserir_compra()
Attribute inserir_compra.VB_ProcData.VB_Invoke_Func = " \n14"
'
' inserir_compra Macro
'

'
    Range("B5:I5").Select
    Selection.ListObject.ListRows.Add (1)
    Range("B5").Select
End Sub
