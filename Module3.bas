Attribute VB_Name = "Module3"
Sub 總和()
Attribute 總和.VB_ProcData.VB_Invoke_Func = "s\n14"
'
' 總和 巨集
'
' 快速鍵: Ctrl+s
'
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C[-3]:R[413]C[-3])"
    Range("E1").Select
End Sub
Sub 平均()
Attribute 平均.VB_ProcData.VB_Invoke_Func = "a\n14"
'
' 平均 巨集
'
' 快速鍵: Ctrl+a
'
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[1]C[-5]:R[413]C[-5])"
    Range("G1").Select
    Application.Goto Reference:="平均"
    Range("G1").Select
    ActiveWorkbook.Save
    ActiveWorkbook.Save
    Range("D5").Select
    ActiveCell.FormulaR1C1 = ""
    Range("E1").Select
    ActiveCell.FormulaR1C1 = ""
    Range("G2").Select
    ActiveCell.FormulaR1C1 = ""
    Range("E1").Select
    ActiveCell.FormulaR1C1 = ""
    Range("H9").Select
    Application.WindowState = xlMinimized
    Application.WindowState = xlNormal
End Sub
