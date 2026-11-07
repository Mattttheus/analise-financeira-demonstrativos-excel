Attribute VB_Name = "Módulo2"
Sub inserir_produto()
Attribute inserir_produto.VB_ProcData.VB_Invoke_Func = " \n14"
'
' inserir_produto Macro
'

'
    Range("B5:M5").Select
    Selection.ListObject.ListRows.Add (1)
    Range("B6:M6").Select
    Selection.Copy
    Range("B5").Select
    ActiveSheet.Paste
    Range("D5").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("F5").Select
    Selection.ClearContents
    Range("I5").Select
    Selection.ClearContents
    Range("J5").Select
    Selection.ClearContents
    Range("D5").Select
End Sub
Sub nova_venda()
Attribute nova_venda.VB_ProcData.VB_Invoke_Func = " \n14"
'
' nova_venda Macro
'

'
    Range("B5:M5").Select
    Selection.ListObject.ListRows.Add (1)
    Range("C6").Select
    Selection.Copy
    Range("C5").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=TODAY()"
    Range("C5").Select
    Selection.Copy
    Range("B5").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[1]C+1"
    Range("B5:C5").Select
    Selection.Copy
    Range("B5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D5").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = ""
    Range("D5").Select
End Sub
