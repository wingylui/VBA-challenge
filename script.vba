Sub Tablecreation(n, s, e)
' Creating titles
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("I1:L1").Font.Bold = True

' Loop for doing calculation of each ticker
    Range("A2").Select
    Do While ActiveCell <> ""
' Check the date for start of the year and copy ticker
        If InStr(1, ActiveCell.Range("B1").Value, s) > 0 Then
            ActiveCell.Range("I1").Value = ActiveCell.Range("A1").Value
            ActiveCell.Offset(n - 1, 0).Select
' Check the date for the end of the year
            If InStr(1, ActiveCell.Range("B1").Value, e) > 0 Then
' Calculate yearly change, % change and Total stock volume
                ActiveCell.Offset(-n + 1, 9).Select
                ActiveCell.FormulaR1C1 = "=R[250]C[-4]-RC[-7]"
                ActiveCell.Range("B1").FormulaR1C1 = "=RC[-1]/RC[-8]"
                ActiveCell.Range("B1").NumberFormat = "0.00%"
                ActiveCell.Range("C1").FormulaR1C1 = "=SUM(RC[-5]:R[250]C[-5])"
                ActiveCell.Offset(n, -9).Select
            Else:
                ActiveCell.Offset(-n, 8).Select
                ActiveCell.Range("A1").Value = "Error"
                ActiveCell.Offset(n + 1, -9).Select
            End If
        Else:
            ActiveCell.Range("I1").Value = "Error"
            ActiveCell.Offset(n, 0).Select
        End If
    Loop
 
'Removing any blank rows
    Columns("I:L").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.Delete Shift:=xlUp
End Sub

 
Sub MaxMinPercentageAndTotalVolume()
'Apply conditional formatting rules
    Dim SetRange As Range
        Set SetRange = Range("J:J")
    Dim Title As Range
        Set Title = Range("J1")
    SetRange.FormatConditions.Delete
    Title.FormatConditions.Add Type:=xlExpression, Formula1:="=TRUE"
        Title.FormatConditions(1).Interior.Color = RGB(255, 255, 255)
    SetRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
        SetRange.FormatConditions(2).Interior.Color = RGB(238, 99, 99)
    SetRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
        SetRange.FormatConditions(3).Interior.Color = RGB(143, 188, 143)
 
' Insert Title
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
 
' Calculation of Max and Min value
    Range("Q2").FormulaR1C1 = "=MAX(C[-6])"
    Range("Q3").FormulaR1C1 = "=MIN(C[-6])"
    Range("Q4").FormulaR1C1 = "=MAX(C[-5])"
 
' Finding the correlated Ticker
    Range("P2").FormulaR1C1 = "=INDEX(R2C9:R3001C9,MATCH(RC[1],R2C[-5]:R3001C11,0))"
    Range("P3").FormulaR1C1 = "=INDEX(R2C9:R3001C9,MATCH(RC[1],R2C[-5]:R3001C11,0))"
    Range("P4").FormulaR1C1 = "=INDEX(R2C9:R3001C9,MATCH(RC[1],R2C[-4]:R3001C12,0))"

End Sub
 
 
 
 
Sub main()
 
' Calculate 2018
    Worksheets("2018").Activate
    Call Tablecreation(251, 20180102, 20181231)
    Call MaxMinPercentageAndTotalVolume
 
' Calculate 2019
    Worksheets("2019").Activate
    Call Tablecreation(252, 20190102, 20191231)
    Call MaxMinPercentageAndTotalVolume
 
' Calculate 2020
    Worksheets("2020").Activate
    Call Tablecreation(253, 20200102, 20201231)
    Call MaxMinPercentageAndTotalVolume

End Sub