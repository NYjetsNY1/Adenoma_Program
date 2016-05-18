Attribute VB_Name = "Module4"
Sub Patient_Total()
Attribute Patient_Total.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Written By Benjamin Sklar
' Patient_Total Macro
'

'
    Application.DisplayAlerts = False
    ActiveWorkbook.CheckCompatibility = False
    
        myFile = Application.GetSaveAsFilename(FileFilter:= _
"Text Files (*.txt), *.txt", Title:="Save Location")

    Range("G2").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-6], ""C, R"")"
    Range("G3").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-6], ""G, B"")"
    Range("G4").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-6], ""G, G"")"
    Range("G5").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-6], ""N, N"")"
    Range("G6").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-6], ""P, R"")"
    Range("G7").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-6], ""S, B"")"
    Range("G8").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-6], ""W, S"")"
    
    Cols = Cells(Rows.Count, "G").End(xlUp).Row
    Open myFile For Output As #1
    For i = 2 To Cols Step 1
        cellValue = Range("G" & i).Text
        Print #1, cellValue
    Next i
    Close #1
    
End Sub
