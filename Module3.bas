Attribute VB_Name = "Module3"
Sub �`�M()
Attribute �`�M.VB_ProcData.VB_Invoke_Func = "s\n14"
'
' �`�M ����
'
' �ֳt��: Ctrl+s
'
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C[-3]:R[413]C[-3])"
    Range("E1").Select
End Sub
Sub ����()
Attribute ����.VB_ProcData.VB_Invoke_Func = "a\n14"
'
' ���� ����
'
' �ֳt��: Ctrl+a
'
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[1]C[-5]:R[413]C[-5])"
    Range("G1").Select
    Application.Goto Reference:="����"
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
