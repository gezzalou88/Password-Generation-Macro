Attribute VB_Name = "Module1"
Sub CreateIDsMain()
    
    Dim i As Long
    
    Dim Sample As String
    
    Dim Counter As Long
    
    Dim Counter2 As Long
    
    Counter = 2
    
    Do While Sheets("Controls").Cells(Counter, "B").Value <> ""
        For i = 1 To Sheets("Controls").Cells(Counter, "A").Value + 1
            Sample = ""
            If i = 1 Then
                Sample = Sample + Sheets("Controls").Cells(1, "B").Value + ";" + CStr(Sheets("Controls").Cells(1, "C").Value)
                
                Counter2 = 4
                
                Do While Sheets("Controls").Cells(1, Counter2).Value <> ""
                    Sample = Sample + ";" + CStr(Sheets("Controls").Cells(1, Counter2).Value)
                    Counter2 = Counter2 + 1
                Loop
            Else
                Sample = Sample + Sheets("Controls").Cells(Counter, "B").Value + CStr(Int(Rnd() * 899999) + 100000) + ";" + CStr(Sheets("Controls").Cells(Counter, "C").Value)
                
                Counter2 = 4
                
                Do While Sheets("Controls").Cells(Counter, Counter2).Value <> ""
                    Sample = Sample + ";" + CStr(Sheets("Controls").Cells(Counter, Counter2).Value)
                    Counter2 = Counter2 + 1
                Loop
            End If
            Sheets("Main Passwords").Cells(i, Counter - 1).Value = Sample
        Next i
        
        Counter = Counter + 1
    Loop
    
    Call CreateIDsPRFs
End Sub

Sub CreateIDsPRFs()
    
    Dim i As Long
    
    Dim j As Long
    
    Dim k As Long
    
    Dim Sample As String
    
    Dim Counter As Long
    
    Dim Counter2 As Long
    
    For j = 1 To 20
        
        Counter = 2
    
        If j = 1 Then
            Do While Sheets("Controls").Cells(Counter, "B").Value <> ""
                For i = 1 To Sheets("Controls").Cells(Counter, "A").Value + 1
                    Sample = ""
                    If i = 1 Then
                    
                        Counter2 = 4
                        
                        Sample = Sample + Sheets("Controls").Cells(1, "B").Value + ";" + CStr(Sheets("Controls").Cells(1, "C").Value)
                        
                        Do While Sheets("Controls").Cells(1, Counter2).Value <> ""
                            Sample = Sample + ";" + CStr(Sheets("Controls").Cells(1, Counter2).Value)
                            Counter2 = Counter2 + 1
                        Loop
                    Else
                        Counter2 = 4
                        
                        If Left(Sheets("Main Passwords").Cells(i, Counter - 1).Value, 8) <> "" Then
                            If j < 10 Then
                                Sample = Sample + Left(Sheets("Main Passwords").Cells(i, Counter - 1).Value, 8) + "_0" + CStr(j) + ";" + CStr(Sheets("Controls").Cells(Counter, "C").Value)
                            Else
                                Sample = Sample + Left(Sheets("Main Passwords").Cells(i, Counter - 1).Value, 8) + "_" + CStr(j) + ";" + CStr(Sheets("Controls").Cells(Counter, "C").Value)
                            End If
                            
                            Do While Sheets("Controls").Cells(Counter, Counter2).Value <> ""
                                Sample = Sample + ";" + CStr(Sheets("Controls").Cells(Counter, Counter2).Value)
                                Counter2 = Counter2 + 1
                            Loop
                        End If
                    End If
                    
                    Sheets("PRF Passwords").Cells(i, Counter - 1).Value = Sample
                Next i
                
                Counter = Counter + 1
            Loop
        Else
            Do While Sheets("Controls").Cells(Counter, "B").Value <> ""
                For i = 2 + (Sheets("Controls").Cells(Counter, "A").Value * (j - 1)) To 1 + (Sheets("Controls").Cells(Counter, "A").Value * j)
                    Sample = ""
                    
                    Counter2 = 4
                        
                    k = i - (Sheets("Controls").Cells(Counter, "A").Value * (j - 1))
                        
                    If Left(Sheets("Main Passwords").Cells(k, Counter - 1).Value, 8) <> "" Then
                        If j < 10 Then
                            Sample = Sample + Left(Sheets("Main Passwords").Cells(k, Counter - 1).Value, 8) + "_0" + CStr(j) + ";" + CStr(Sheets("Controls").Cells(Counter, "C").Value)
                        Else
                            Sample = Sample + Left(Sheets("Main Passwords").Cells(k, Counter - 1).Value, 8) + "_" + CStr(j) + ";" + CStr(Sheets("Controls").Cells(Counter, "C").Value)
                        End If
                        
                        Do While Sheets("Controls").Cells(Counter, Counter2).Value <> ""
                            Sample = Sample + ";" + CStr(Sheets("Controls").Cells(Counter, Counter2).Value)
                            Counter2 = Counter2 + 1
                        Loop
                    End If
                    
                    Sheets("PRF Passwords").Cells(i, Counter - 1).Value = Sample
                Next i
                
                Counter = Counter + 1
            Loop
        End If
    Next j
    
    Call DeDupeCols
End Sub

Sub DeDupeCols()

Dim Counter As Long

Counter = 2

Do While Sheets("Controls").Cells(Counter, "B").Value <> ""
    Sheets("Main Passwords").Range(Sheets("Main Passwords").Cells(2, Counter - 1), Sheets("Main Passwords").Cells(Sheets("Controls").Cells(Counter, "A").Value + 1, Counter - 1)).RemoveDuplicates Columns:=1
    Sheets("PRF Passwords").Range(Sheets("PRF Passwords").Cells(2, Counter - 1), Sheets("PRF Passwords").Cells(((Sheets("Controls").Cells(Counter, "A").Value + 1) + (Sheets("Controls").Cells(Counter, "A").Value * 19)), Counter - 1)).RemoveDuplicates Columns:=1
    Counter = Counter + 1
Loop

MsgBox "Finished Live ID run"


End Sub

Sub CreateTestsMain()
    
    Dim Counter As Long
    
    Dim Counter2 As Long
    
    Counter = 2
    
    Do While Sheets("Controls").Cells(Counter, "B").Value <> ""
        For i = 1 To 1000
            Sample = ""
            If i = 1 Then
                Sample = Sample + Sheets("Controls").Cells(1, "B").Value + ";" + CStr(Sheets("Controls").Cells(1, "C").Value)
                
                Counter2 = 4
                
                Do While Sheets("Controls").Cells(1, Counter2).Value <> ""
                    Sample = Sample + ";" + CStr(Sheets("Controls").Cells(1, Counter2).Value)
                    Counter2 = Counter2 + 1
                Loop
            Else
                If i < 11 Then
                    Sample = Sample + Sheets("Controls").Cells(Counter, "B").Value + "Test00" + CStr(i - 1) + ";" + CStr(Sheets("Controls").Cells(Counter, "C").Value)
                ElseIf i > 10 And i < 101 Then
                    Sample = Sample + Sheets("Controls").Cells(Counter, "B").Value + "Test0" + CStr(i - 1) + ";" + CStr(Sheets("Controls").Cells(Counter, "C").Value)
                Else
                    Sample = Sample + Sheets("Controls").Cells(Counter, "B").Value + "Test" + CStr(i - 1) + ";" + CStr(Sheets("Controls").Cells(Counter, "C").Value)
                End If
                
                Counter2 = 4
                
                Do While Sheets("Controls").Cells(Counter, Counter2).Value <> ""
                    Sample = Sample + ";" + CStr(Sheets("Controls").Cells(Counter, Counter2).Value)
                    Counter2 = Counter2 + 1
                Loop
            End If
            Sheets("Test Passwords").Cells(i, Counter - 1).Value = Sample
        Next i
        
        Counter = Counter + 1
    Loop
    
    Call CreateTestsPRFs
End Sub

Sub CreateTestsPRFs()
    
    Dim i As Long
    
    Dim j As Long
    
    Dim k As Long
    
    Dim Sample As String
    
    Dim Counter As Long
    
    Dim Counter2 As Long
    
    For j = 1 To 20
        
        Counter = 2
    
        If j = 1 Then
            Do While Sheets("Controls").Cells(Counter, "B").Value <> ""
                For i = 1 To 1000
                    Sample = ""
                    
                    If i = 1 Then
                    
                        Counter2 = 4
                        
                        Sample = Sample + Sheets("Controls").Cells(1, "B").Value + ";" + CStr(Sheets("Controls").Cells(1, "C").Value)
                        
                        Do While Sheets("Controls").Cells(1, Counter2).Value <> ""
                            Sample = Sample + ";" + CStr(Sheets("Controls").Cells(1, Counter2).Value)
                            Counter2 = Counter2 + 1
                        Loop
                    Else
                        Counter2 = 4
                        
                        If Left(Sheets("Test Passwords").Cells(i, Counter - 1).Value, 8) <> "" Then
                            If j < 10 Then
                                Sample = Sample + Left(Sheets("Test Passwords").Cells(i, Counter - 1).Value, 9) + "_0" + CStr(j) + ";" + CStr(Sheets("Controls").Cells(Counter, "C").Value)
                            Else
                                Sample = Sample + Left(Sheets("Test Passwords").Cells(i, Counter - 1).Value, 9) + "_" + CStr(j) + ";" + CStr(Sheets("Controls").Cells(Counter, "C").Value)
                            End If
                            
                            Do While Sheets("Controls").Cells(Counter, Counter2).Value <> ""
                                Sample = Sample + ";" + CStr(Sheets("Controls").Cells(Counter, Counter2).Value)
                                Counter2 = Counter2 + 1
                            Loop
                        End If
                    End If
                    
                    Sheets("Test PRF Passwords").Cells(i, Counter - 1).Value = Sample
                Next i
                
                Counter = Counter + 1
            Loop
        Else
            Do While Sheets("Controls").Cells(Counter, "B").Value <> ""
                For i = 2 + (999 * (j - 1)) To 1 + (999 * j)
                    Sample = ""
                    
                    Counter2 = 4
                        
                    k = i - (999 * (j - 1))
                        
                    If Left(Sheets("Test Passwords").Cells(k, Counter - 1).Value, 9) <> "" Then
                        If j < 10 Then
                            Sample = Sample + Left(Sheets("Test Passwords").Cells(k, Counter - 1).Value, 9) + "_0" + CStr(j) + ";" + CStr(Sheets("Controls").Cells(Counter, "C").Value)
                        Else
                            Sample = Sample + Left(Sheets("Test Passwords").Cells(k, Counter - 1).Value, 9) + "_" + CStr(j) + ";" + CStr(Sheets("Controls").Cells(Counter, "C").Value)
                        End If
                        
                        Do While Sheets("Controls").Cells(Counter, Counter2).Value <> ""
                            Sample = Sample + ";" + CStr(Sheets("Controls").Cells(Counter, Counter2).Value)
                            Counter2 = Counter2 + 1
                        Loop
                    End If
                    
                    Sheets("Test PRF Passwords").Cells(i, Counter - 1).Value = Sample
                Next i
                
                Counter = Counter + 1
            Loop
        End If
    Next j
    
    MsgBox "Finished Test ID run"
    
End Sub

Sub ExportIDs()

    Dim WS As Worksheet

    Dim i As Long
    
    Dim j As Long
    
    Dim Sample As String
    Dim Sample1 As String
    
    Dim SaveFile As String
    
    Dim Counter As Long
        
    Counter = 2
    
    Do While Sheets("Controls").Cells(Counter, "B").Value <> ""
    
        Set WS = Sheets.Add(After:=Sheets(Sheets.Count))
        WS.Name = Sheets("Controls").Cells(Counter, "B").Value
        
        SaveFile = CStr(Sheets("Controls").Cells(37, "C").Value) + "\" + CStr(Sheets("Controls").Cells(36, "C").Value) + " " + CStr(Sheets("Controls").Cells(Counter, "B").Value) + " Live IDs.xlsx"
    
        For i = 1 To Sheets("Controls").Cells(Counter, "A").Value + 1
            Sample = ""
            Sample1 = ""
            
            If i = 1 Then
               Sample = Sample + "Live ID"
               Sample1 = Sample1 + "Live Link"
            Else
                Sample = Sample + Left(Sheets("Main Passwords").Cells(i, Counter - 1).Value, 8)
                Sample1 = Sample1 + CStr(Sheets("Controls").Cells(38, "C").Value) + Left(Sheets("Main Passwords").Cells(i, Counter - 1).Value, 8)
            End If
                    
            Sheets(Sheets("Controls").Cells(Counter, "B").Value).Cells(i, "A").Value = Sample
            Sheets(Sheets("Controls").Cells(Counter, "B").Value).Cells(i, "B").Value = Sample1
        Next i
        
        For i = 749 To 999
            Sample = ""
            Sample1 = ""
            
            j = i - 748
            
            If i = 749 Then
                Sample = Sample + "Test ID"
                Sample1 = Sample1 + "Test Link"
            Else
                Sample = Sample + Sheets("Controls").Cells(Counter, "B").Value + "Test" + CStr(i - 1)
                Sample1 = Sample1 + CStr(Sheets("Controls").Cells(38, "C").Value) + Sheets("Controls").Cells(Counter, "B").Value + "Test" + CStr(i - 1)
            End If
            Sheets(Sheets("Controls").Cells(Counter, "B").Value).Cells(j, "C").Value = Sample
            Sheets(Sheets("Controls").Cells(Counter, "B").Value).Cells(j, "D").Value = Sample1
        Next i
        
        Sheets(Sheets("Controls").Cells(Counter, "B").Value).Cells.EntireColumn.AutoFit
        
        Sheets(Sheets("Controls").Cells(Counter, "B").Value).Move

        ActiveWorkbook.SaveAs Filename:=SaveFile _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        
        MsgBox (SaveFile)
        
        ActiveWindow.Close
                
        Counter = Counter + 1
    Loop

End Sub

Sub ClearContent()

Sheets("Main Passwords").Cells.ClearContents
Sheets("PRF Passwords").Cells.ClearContents
Sheets("Test Passwords").Cells.ClearContents
Sheets("Test PRF Passwords").Cells.ClearContents

End Sub

Sub ExportTestSample()

    Dim WS1 As Worksheet
    Dim i As Long
    Dim j As Long
    Dim Sample1 As String
    Dim SaveFile1 As String
    Dim Counter As Long
    Dim Counter1 As Long
    Dim StartRow As Long
        
    Counter = 2
    Counter1 = 1
    
    Set WS1 = Sheets.Add(After:=Sheets(Sheets.Count))
    WS1.Name = "Test IDs"
    SaveFile1 = "I:\Administration\Data Services\Survey Programming\Templates\Live and Test Sample\Test IDs.csv"
    
    Do While Sheets("Controls").Cells(Counter, "B").Value <> ""
        If Counter1 = 1 Then
            For i = 1 To 1000
                Sample1 = ""
                Sample1 = Sample1 + Sheets("Test Passwords").Cells(i, "A").Value
                Sheets("Test IDs").Cells(i, "A").Value = Sample1
            Next i
        Else
            For i = 2 To 1000
                StartRow = (i - (Counter1 - 1)) + ((Counter1 - 1) * 1000)
                Sample1 = ""
                Sample1 = Sample1 + Sheets("Test Passwords").Cells(i, Counter1).Value
                Sheets("Test IDs").Cells(StartRow, "A").Value = Sample1
            Next i
        End If
        
        Counter = Counter + 1
        Counter1 = Counter1 + 1
    Loop
        
    Sheets("Test IDs").Cells.EntireColumn.AutoFit
    Sheets("Test IDs").Move

    ActiveWorkbook.SaveAs Filename:=SaveFile1 _
        , FileFormat:=xlCSV, CreateBackup:=False

    MsgBox (SaveFile1)
    ActiveWindow.Close
    
    Counter = 2
    Counter1 = 1
    
    Set WS1 = Sheets.Add(After:=Sheets(Sheets.Count))
    WS1.Name = "Test PRF IDs"
    SaveFile1 = "I:\Administration\Data Services\Survey Programming\Templates\Live and Test Sample\Test PRF IDs.csv"
    
    Do While Sheets("Controls").Cells(Counter, "B").Value <> ""
        If Counter1 = 1 Then
            For i = 1 To 19981
                Sample1 = ""
                Sample1 = Sample1 + Sheets("Test PRF Passwords").Cells(i, "A").Value
                Sheets("Test PRF IDs").Cells(i, "A").Value = Sample1
            Next i
        Else
            For i = 2 To 19981
                StartRow = (i - (Counter1 - 1)) + ((Counter1 - 1) * 19981)
                Sample1 = ""
                Sample1 = Sample1 + Sheets("Test PRF Passwords").Cells(i, Counter1).Value
                Sheets("Test PRF IDs").Cells(StartRow, "A").Value = Sample1
            Next i
        End If
        
        Counter = Counter + 1
        Counter1 = Counter1 + 1
    Loop
        
    Sheets("Test PRF IDs").Cells.EntireColumn.AutoFit
    Sheets("Test PRF IDs").Move

    ActiveWorkbook.SaveAs Filename:=SaveFile1 _
        , FileFormat:=xlCSV, CreateBackup:=False

    MsgBox (SaveFile1)
    ActiveWindow.Close
    
    
    
End Sub

Sub ExportLiveSample()

    Dim WS1 As Worksheet
    Dim i As Long
    Dim j As Long
    Dim Sample1 As String
    Dim SaveFile1 As String
    Dim Counter As Long
    Dim Counter1 As Long
    Dim StartRow As Long
        
    Counter = 2
    Counter1 = 1
    
    Set WS1 = Sheets.Add(After:=Sheets(Sheets.Count))
    WS1.Name = "Live IDs"
    SaveFile1 = "I:\Administration\Data Services\Survey Programming\Templates\Live and Test Sample\Live IDs.csv"
    
    Do While Sheets("Controls").Cells(Counter, "B").Value <> ""
        If Counter1 = 1 Then
            For i = 1 To (Sheets("Controls").Cells(Counter, "A").Value + 1)
                Sample1 = ""
                Sample1 = Sample1 + Sheets("Main Passwords").Cells(i, "A").Value
                Sheets("Live IDs").Cells(i, "A").Value = Sample1
            Next i
        Else
            For i = 2 To (Sheets("Controls").Cells(Counter, "A").Value + 1)
                StartRow = (i - (Counter1 - 1)) + ((Counter1 - 1) * (Sheets("Controls").Cells(Counter, "A").Value + 1))
                Sample1 = ""
                Sample1 = Sample1 + Sheets("Main Passwords").Cells(i, Counter1).Value
                Sheets("Live IDs").Cells(StartRow, "A").Value = Sample1
            Next i
        End If
        
        Counter = Counter + 1
        Counter1 = Counter1 + 1
    Loop
        
    Sheets("Live IDs").Cells.EntireColumn.AutoFit
    Sheets("Live IDs").Move
    
    Sheets("Live IDs").Sort.SortFields.Clear
    Sheets("Live IDs").Sort.SortFields.Add Key:=Range( _
        "A2:A18596"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With Sheets("Live IDs").Sort
        .SetRange Range("A1:A18596")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ActiveWorkbook.SaveAs Filename:=SaveFile1 _
        , FileFormat:=xlCSV, CreateBackup:=False

    MsgBox (SaveFile1)
    ActiveWindow.Close
    
    Counter = 2
    Counter1 = 1
    
    Set WS1 = Sheets.Add(After:=Sheets(Sheets.Count))
    WS1.Name = "Live PRF IDs"
    SaveFile1 = "I:\Administration\Data Services\Survey Programming\Templates\Live and Test Sample\Live PRF IDs.csv"
    
    Do While Sheets("Controls").Cells(Counter, "B").Value <> ""
        If Counter1 = 1 Then
            For i = 1 To ((Sheets("Controls").Cells(Counter, "A").Value + 1) * 20)
                Sample1 = ""
                Sample1 = Sample1 + Sheets("PRF Passwords").Cells(i, "A").Value
                Sheets("Live PRF IDs").Cells(i, "A").Value = Sample1
            Next i
        Else
            For i = 2 To ((Sheets("Controls").Cells(Counter, "A").Value + 1) * 20)
                StartRow = (i - (Counter1 - 1)) + ((Counter1 - 1) * ((Sheets("Controls").Cells(Counter, "A").Value + 1) * 20))
                Sample1 = ""
                Sample1 = Sample1 + Sheets("PRF Passwords").Cells(i, Counter1).Value
                Sheets("Live PRF IDs").Cells(StartRow, "A").Value = Sample1
            Next i
        End If
        
        Counter = Counter + 1
        Counter1 = Counter1 + 1
    Loop
        
    Sheets("Live PRF IDs").Cells.EntireColumn.AutoFit
    Sheets("Live PRF IDs").Move
    
    Sheets("Live PRF IDs").Sort.SortFields.Clear
    Sheets("Live PRF IDs").Sort.SortFields.Add Key:=Range( _
        "A2:A18596"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With Sheets("Live PRF IDs").Sort
        .SetRange Range("A1:A18596")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ActiveWorkbook.SaveAs Filename:=SaveFile1 _
        , FileFormat:=xlCSV, CreateBackup:=False

    MsgBox (SaveFile1)
    ActiveWindow.Close
    
    
    
End Sub

