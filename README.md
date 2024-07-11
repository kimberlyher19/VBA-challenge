# VBA-challenge
original code was put into chatgpt to help with coding and running through all ws
functions were used in excel and converted to vba codes:

Sub Min_Max()
'
' Min_Max Macro
'

'
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "Min"
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "Max"
    Range("M2").Select
    ActiveCell.FormulaR1C1 = _
        "=MINIFS(C[-10], C[-11], MINIFS(C[-11], C[-12], RC[-4]), C[-12], RC[-4])"
    Range("M2").Select
    Selection.AutoFill Destination:=Range("M2:M1501")
    Range("M2:M1501").Select
    Range("N2").Select
    ActiveSheet.Paste
    Selection.AutoFill Destination:=Range("N2:N1501")
    Range("N2:N1501").Select

End Sub

Sub Quarterly_Change()
'
' Quarterly_Change Macro
'

'
    ActiveCell.FormulaR1C1 = "=RC[4]-RC[3]"
    Range("J2").Select
    Selection.AutoFill Destination:=Range("J2:J1501")
    Range("J2:J1501").Select
End Sub
Sub PercentChange()
'
' PercentChange Macro
'

'
    ActiveCell.FormulaR1C1 = "=((RC[3]-RC[2])/RC[2])"
    Range("K2").Select
    Selection.NumberFormat = "0.00%"
    Selection.AutoFill Destination:=Range("K2:K1501")
    Range("K2:K1501").Select
End Sub

Sub Volume()
'
' Volume Macro
'

'
    ActiveCell.FormulaR1C1 = "=SUMIFS(C[-5], C[-11], RC[-3])"
    Range("L2").Select
    Selection.AutoFill Destination:=Range("L2:L1501")
    Range("L2:L1501").Select
   
End Sub

ActiveCell.FormulaR1C1 = "=MAX(C[-7])"
    Range("R3").Select
    ActiveCell.FormulaR1C1 = "=MIN(C[-7])"
    Range("R4").Select
    ActiveCell.FormulaR1C1 = "=MAX(C[-6])"
    Range("R5").Select
