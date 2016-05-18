Attribute VB_Name = "Module3"
Sub Adenoma_Duplicates()
Attribute Adenoma_Duplicates.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Written By Benjamin Sklar
' Adenoma_Duplicates Macro
'

'
    Application.DisplayAlerts = False
    ActiveWorkbook.CheckCompatibility = False
    ActiveSheet.Range("$A:$I").RemoveDuplicates Columns:=3, Header:=xlYes
    myFile = Application.GetSaveAsFilename(FileFilter:= _
"Text Files (*.txt), *.txt", Title:="Save Location")

    
    LastRow = Cells(Rows.Count, "B").End(xlUp).Row
    For i = LastRow To 2 Step -1
        Range("K" & i).FormulaR1C1 = "=ROUNDDOWN((RC[-5]-RC[-7])/365,0)"
    Next i
    
    For i = LastRow To 2 Step -1
    If Range("K" & i).Value < 50 Then
        Range("K" & i).EntireRow.Delete
    End If
    Next i
    

    Range("L2").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-10], ""C, R"")"
    Range("L3").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-10], ""G, B"")"
    Range("L4").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-10], ""G, G"")"
    Range("L5").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-10], ""N, N"")"
    Range("L6").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-10], ""P, R"")"
    Range("L7").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-10], ""S, B"")"
    Range("L8").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-10], ""W, S"")"
    Range("L9").Select
    
    
    Cols = Cells(Rows.Count, "L").End(xlUp).Row
    Open myFile For Output As #1
    For i = 2 To Cols Step 1
        cellValue = Range("L" & i).Text
        Print #1, cellValue
    Next i
    Close #1
    

End Sub

