Attribute VB_Name = "Module2"
Sub calAVG2()
Attribute calAVG2.VB_Description = "�������D�n�Ω�p�⥭����"
Attribute calAVG2.VB_ProcData.VB_Invoke_Func = "w\n14"
'
' calAVG2 ����
' �������D�n�Ω�p�⥭����
'
' �ֳt��: Ctrl+w
'
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C[-3]:R[413]C[-3])"
    Range("E2").Select
    ActiveWindow.SmallScroll Down:=-24
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C[-3]:R[413]C[-3])"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[1]C[-5]:R[413]C[-5])"
    Range("G2").Select
End Sub
Sub calavg3()
Attribute calavg3.VB_Description = "�p��~~~~�ЦP�Ǹԭz�����\��"
Attribute calavg3.VB_ProcData.VB_Invoke_Func = "b\n14"
'
' calavg3 ����
' �p��~~~~�ЦP�Ǹԭz�����\��
'
' �ֳt��: Ctrl+b
'
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C[-3]:R[413]C[-3])"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[1]C[-5]:R[413]C[-5])"
    Range("G2").Select
End Sub
