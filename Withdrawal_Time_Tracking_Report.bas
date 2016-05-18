Attribute VB_Name = "Module2"
Sub Time_Tracking()
Attribute Time_Tracking.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Written By Benjamin Sklar
' Macro2 Macro
'

'
    Application.DisplayAlerts = False
    ActiveWorkbook.CheckCompatibility = False
    myFile = Application.GetSaveAsFilename(FileFilter:= _
"Text Files (*.txt), *.txt", Title:="Save Location")
    LastRow = Cells(Rows.Count, "B").End(xlUp).Row


    For i = LastRow To 1 Step -1
        If Range("B" & i).Value <> "Colonoscopy" And Range("B" & i).Value <> "Colonoscopy, Upper GI endoscopy" And Range("B" & i).Value <> "Upper GI endoscopy, Colonoscopy" Then
            Range("B" & i).EntireRow.Delete
        End If
        If Range("N" & i).Value = "-" Then
            Range("N" & i).EntireRow.Delete
        End If
        If Range("M" & i).Value = "-" Then
            Range("M" & i).EntireRow.Delete
        End If
    Next i
    

    Range("S1").Select
    ActiveCell.FormulaR1C1 = "=(RC[-6]-RC[-5])"
    Selection.AutoFill Destination:=Range("S1:S" & LastRow)
    For i = LastRow To 1 Step -1
    If Range("S" & i).Value <= 0 Then
        Range("S" & i).ClearContents
    End If
    Next i
    

    Range("T1").Select
    ActiveCell.FormulaR1C1 = "=(RC[-4]-RC[-6])"
    Range("T1").Select
    Selection.AutoFill Destination:=Range("T:T")
    Range("T:T").Select
    Range("U1").Select
    ActiveCell.FormulaR1C1 = "=AVERAGEIF(C[-16],""R C"",C[-2]:C[-1])"
    Range("U2").Select
    ActiveCell.FormulaR1C1 = "=AVERAGEIF(C[-16],""B G"",C[-2]:C[-1])"
    Range("U3").Select
    ActiveCell.FormulaR1C1 = "=AVERAGEIF(C[-16],""G G"",C[-2]:C[-1])"
    Range("U4").Select
    ActiveCell.FormulaR1C1 = "=AVERAGEIF(C[-16],""N N"",C[-2]:C[-1])"
    Range("U5").Select
    ActiveCell.FormulaR1C1 = "=AVERAGEIF(C[-16],"" P"",C[-2]:C[-1])"
    Range("U6").Select
    ActiveCell.FormulaR1C1 = "=AVERAGEIF(C[-16],""B S"",C[-2]:C[-1])"
    Range("U7").Select
    ActiveCell.FormulaR1C1 = "=AVERAGEIF(C[-16],""S W"",C[-2]:C[-1])"
    Columns("U:U").Select
    Selection.NumberFormat = "mm:ss"
    

    Cols = Cells(Rows.Count, "U").End(xlUp).Row
    Open myFile For Output As #1
    For i = 1 To Cols Step 1
        cellValue = Range("U" & i).Text
        Print #1, cellValue
    Next i
    Close #1
    
    
    
End Sub
