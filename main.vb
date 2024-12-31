Sub RUN_COMPARE()
    Call RoundInPlace
    Call RemoveSpacesAllCells
    Call CopyFilteredDataToMultipleSheets
    Call ClearDataAndFormat
    Call MoveDataToEvenRowsAndClearSource
    Call FindAndFillDataWithDetails
    Call ColorRows
    Call CompareAndHighlightDifferences
    Call DeleteRowsAndCleanColumnP
    Call AddPart
    Call DeletePart
End Sub


Sub RoundInPlace()
    Dim cell As Range
    For Each cell In Selection
        If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            cell.Value = Application.Round(cell.Value, 2)
        End If
    Next cell
End Sub

Sub RemoveSpacesAllCells()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim colRanges As Variant
    Dim colRange As Variant
    Dim lastRow As Long
    Dim startCol As String, endCol As String

    Set ws = ActiveSheet

    colRanges = Array("A:Y", "AN:BR", "CK:DJ")

    Application.ScreenUpdating = False

    For Each colRange In colRanges
        startCol = Split(colRange, ":")(0)
        endCol = Split(colRange, ":")(1)

        Dim col As Range
        For Each col In ws.Range(startCol & "1:" & endCol & ws.Rows.Count).Columns
            lastRow = ws.Cells(ws.Rows.Count, col.Column).End(xlUp).Row

            Set rng = ws.Range(ws.Cells(1, col.Column), ws.Cells(lastRow, col.Column))

            For Each cell In rng
                If Not IsEmpty(cell.Value) Then
                    If IsNumeric(cell.Value) Then
                        cell.Value = cell.Value * 1 
                    Else
                        cell.Value = Trim(cell.Value) 
                    End If
                End If
            Next cell
        Next col
    Next colRange

    Application.ScreenUpdating = True

End Sub


Sub CopyFilteredDataToMultipleSheets()
    Dim wsSource As Worksheet
    Dim wsTargetPartlist As Worksheet
    Dim wsTargetMartlist As Worksheet
    Dim lastRow As Long
    Dim targetRowPartlist As Long
    Dim targetRowMartlist As Long
    Dim i As Long, j As Long
    Dim uniqueFound As Boolean
    Dim colStart As Integer, colEnd As Integer
    Dim codePart As String
    Dim matchFound As Boolean

    Set wsSource = ThisWorkbook.Sheets("ALL INPUT") ' Sheet sumber
    Set wsTargetPartlist = ThisWorkbook.Sheets("RESULT PARTLIST") ' Sheet tujuan PARTLIST
    
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    
    targetRowPartlist = 2
    
    colStart = 29
    colEnd = 34
    
    skipText = Array("Block", "Block Code", "Panel", "Code-P-SP-Part")

    For i = 2 To lastRow
        uniqueFound = False
        matchFound = False
        
        If wsSource.Cells(i, 28).Value = "Duplikat" Then 
            For j = colStart To colEnd
                If wsSource.Cells(i, j).Value = "Unik" Then
                    uniqueFound = True
                    Exit For
                End If
            Next j
            
            If uniqueFound Then
                codePart = wsSource.Cells(i, 14).Value 
                Dim textFound As Boolean
                textFound = False
                For Each Skip In skipText
                    If InStr(1, codePart, Skip, vbTextCompare) > 0 Then
                        textFound = True
                        Exit For
                    End If
                Next Skip

                If textFound Then GoTo NextRowPartlist
                
                For j = 2 To lastRow
                    If wsSource.Cells(j, 1).Value = codePart Then 
                        matchFound = True
                        Exit For
                    End If
                Next j
                
                If matchFound Then
                    wsTargetPartlist.Range(wsTargetPartlist.Cells(targetRowPartlist, 1), wsTargetPartlist.Cells(targetRowPartlist, 34)).Value = _
                        wsSource.Range(wsSource.Cells(i, 1), wsSource.Cells(i, 34)).Value
                    targetRowPartlist = targetRowPartlist + 1
                End If
            End If
        End If
NextRowPartlist:
    Next i
    
    Set wsTargetMartlist = ThisWorkbook.Sheets("RESULT MARTLIST")
    
    targetRowMartlist = 2
    
    colStart = 74
    colEnd = 85

    For i = 2 To lastRow
        uniqueFound = False
        matchFound = False
        
        If wsSource.Cells(i, 73).Value = "Duplikat" Then 
            For j = colStart To colEnd
                If wsSource.Cells(i, j).Value = "Unik" Then
                    uniqueFound = True
                    Exit For
                End If
            Next j
            
            If uniqueFound Then
                codePart = wsSource.Cells(i, 57).Value 
                For j = 2 To lastRow
                    If wsSource.Cells(j, 41).Value = codePart Then 
                        matchFound = True
                        Exit For
                    End If
                Next j
                
                If matchFound Then
                    wsTargetMartlist.Range(wsTargetMartlist.Cells(targetRowMartlist, 1), wsTargetMartlist.Cells(targetRowMartlist, 46)).Value = _
                        wsSource.Range(wsSource.Cells(i, 40), wsSource.Cells(i, 85)).Value
                    targetRowMartlist = targetRowMartlist + 1
                End If
            End If
        End If
    Next i

    Set wsTargetMartlist = ThisWorkbook.Sheets("RESULT APL")
    
    targetRowMartlist = 2
    
    colStart = 119
    colEnd = 122

    For i = 2 To lastRow
        uniqueFound = False
        matchFound = False
        
        If wsSource.Cells(i, 118).Value = "Duplikat" Then 
            For j = colStart To colEnd
                If wsSource.Cells(i, j).Value = "Unik" Then
                    uniqueFound = True
                    Exit For
                End If
            Next j
            
            If uniqueFound Then
                codePart = wsSource.Cells(i, 104).Value 
                
                For j = 2 To lastRow
                    If wsSource.Cells(j, 101).Value = codePart Then 
                        matchFound = True
                        Exit For
                    End If
                Next j
                
                If matchFound Then
                    wsTargetMartlist.Range(wsTargetMartlist.Cells(targetRowMartlist, 1), wsTargetMartlist.Cells(targetRowMartlist, 46)).Value = _
                        wsSource.Range(wsSource.Cells(i, 89), wsSource.Cells(i, 122)).Value
                    targetRowMartlist = targetRowMartlist + 1
                End If
            End If
        End If
    Next i
End Sub

Sub ClearDataAndFormat()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim ws3 As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim rowToStart As Long
    
    Set ws1 = ThisWorkbook.Sheets("RESULT PARTLIST")
    Set ws2 = ThisWorkbook.Sheets("RESULT MARTLIST")
    Set ws3 = ThisWorkbook.Sheets("RESULT APL")
    
    ws1.Columns("A:L").Clear
    ws1.Columns("AA:AH").Clear
    
    ws2.Columns("A:Q").Clear
    ws2.Columns("AG:AT").Clear

    ws3.Columns("A:M").Clear
    ws3.Columns("X:AT").Clear
    
    lastRow = ws2.Cells(ws2.Rows.Count, "T").End(xlUp).Row
    
    For i = 2 To lastRow
        If ws2.Cells(i, 20).Value = "" Then 
            ws2.Range(ws2.Cells(i, 18), ws2.Cells(i, 31)).Clear
        End If
    Next i

    lastRow = ws2.Cells.Find(What:="*", LookIn:=xlFormulas, _
                              SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

    If lastRow >= 2 Then
        rowToStart = lastRow - 1 
        ws2.Rows(rowToStart & ":" & lastRow).ClearContents
    ElseIf lastRow = 1 Then
        ws2.Rows(lastRow).ClearContents
    End If
End Sub

Sub MoveDataToEvenRowsAndClearSource()
    Dim wsTarget1 As Worksheet
    Dim wsTarget2 As Worksheet
    Dim wsTarget3 As Worksheet
    Dim lastRow1 As Long, lastRow2 As Long, lastRow3 As Long
    Dim targetRow1 As Long, targetRow2 As Long, targetRow3 As Long
    Dim i As Long

    Set wsTarget1 = ThisWorkbook.Sheets("RESULT PARTLIST") 
    Set wsTarget2 = ThisWorkbook.Sheets("RESULT MARTLIST")     
    Set wsTarget3 = ThisWorkbook.Sheets("RESULT APL")    

    lastRow1 = wsTarget1.Cells(wsTarget1.Rows.Count, "N").End(xlUp).Row 
    
    targetRow1 = 2 
    
    For i = 2 To lastRow1
        If wsTarget1.Cells(i, 14).Value <> "" Then 
            wsTarget1.Range(wsTarget1.Cells(i, 14), wsTarget1.Cells(i, 25)).Copy _
                Destination:=wsTarget1.Cells(targetRow1, 1)
            
            targetRow1 = targetRow1 + 2
        End If
    Next i
    
    wsTarget1.Range("N2:Y" & lastRow1).Clear
    
    lastRow2 = wsTarget2.Cells(wsTarget2.Rows.Count, "R").End(xlUp).Row 
    
    targetRow2 = 2 
    
    For i = 2 To lastRow2
        If wsTarget2.Cells(i, 18).Value <> "" Then 
            wsTarget2.Range(wsTarget2.Cells(i, 18), wsTarget2.Cells(i, 31)).Copy _
                Destination:=wsTarget2.Cells(targetRow2, 1)
            targetRow2 = targetRow2 + 2
        End If
    Next i
    
    wsTarget2.Range("R2:AE" & lastRow2).Clear

    lastRow3 = wsTarget3.Cells(wsTarget3.Rows.Count, "P").End(xlUp).Row 
    
    targetRow3 = 2
    
    For i = 2 To lastRow3
        If wsTarget3.Cells(i, 16).Value <> "" Then 
            wsTarget3.Range(wsTarget3.Cells(i, 16), wsTarget3.Cells(i, 21)).Copy _
                Destination:=wsTarget3.Cells(targetRow3, 1)
            
            targetRow3 = targetRow3 + 2
        End If
    Next i
    
    wsTarget3.Range("P2:X" & lastRow3).Clear
    
End Sub

Sub FindAndFillDataWithDetails()
    Dim wsSource As Worksheet
    Dim wsTarget1 As Worksheet
    Dim wsTarget2 As Worksheet
    Dim wsTarget3 As Worksheet
    Dim lastRowSource As Long
    Dim lastRowTarget1 As Long
    Dim lastRowTarget2 As Long
    Dim lastRowTarget3 As Long
    Dim i As Long, j As Long
    Dim codePart As String
    Dim found As Boolean
    Dim targetRow1 As Long, targetRow2 As Long, targetRow3 As Long
    Dim targetRange As Range

    Set wsSource = ThisWorkbook.Sheets("ALL INPUT") 
    Set wsTarget1 = ThisWorkbook.Sheets("RESULT PARTLIST") 
    Set wsTarget2 = ThisWorkbook.Sheets("RESULT MARTLIST") 
    Set wsTarget3 = ThisWorkbook.Sheets("RESULT APL") 

    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row 
    lastRowTarget1 = wsTarget1.Cells(wsTarget1.Rows.Count, "A").End(xlUp).Row 
    lastRowTarget2 = wsTarget2.Cells(wsTarget2.Rows.Count, "A").End(xlUp).Row 
    lastRowTarget3 = wsTarget3.Cells(wsTarget3.Rows.Count, "A").End(xlUp).Row 

    targetRow1 = 3 
    For i = 2 To lastRowTarget1 Step 2 
        codePart = wsTarget1.Cells(i, 1).Value 
        found = False

        For j = 2 To lastRowSource
            If wsSource.Cells(j, 1).Value = codePart Then 
                With wsTarget1
                    .Cells(targetRow1, 1).Value = wsSource.Cells(j, 1).Value
                    .Cells(targetRow1, 2).Resize(1, 11).Value = wsSource.Cells(j, 2).Resize(1, 11).Value
                End With
                found = True
                targetRow1 = targetRow1 + 2
                Exit For
            End If
        Next j

        If Not found Then
            With wsTarget1
                .Cells(targetRow1, 1).Value = "Tidak Ditemukan"
                .Cells(targetRow1, 2).Resize(1, 11).ClearContents
            End With
            targetRow1 = targetRow1 + 2
        End If
    Next i

    Set targetRange = wsTarget1.Range("A3:L" & targetRow1 - 1)
    With targetRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    targetRow2 = 3 
    For i = 2 To lastRowTarget2 Step 2 
        codePart = wsTarget2.Cells(i, 1).Value 
        found = False

        For j = 2 To lastRowSource
            If wsSource.Cells(j, 41).Value = codePart Then 
                With wsTarget2
                    .Cells(targetRow2, 1).Value = wsSource.Cells(j, 41).Value
                    .Cells(targetRow2, 2).Resize(1, 14).Value = wsSource.Cells(j, 42).Resize(1, 14).Value
                End With
                found = True
                targetRow2 = targetRow2 + 2
                Exit For
            End If
        Next j

        If Not found Then
            With wsTarget2
                .Cells(targetRow2, 1).Value = "Tidak Ditemukan"
                .Cells(targetRow2, 2).Resize(1, 14).ClearContents
            End With
            targetRow2 = targetRow2 + 2
        End If
    Next i

    Set targetRange = wsTarget2.Range("A3:N" & targetRow2 - 1)
    With targetRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    targetRow3 = 3 
    For i = 2 To lastRowTarget3 Step 2 
        codePart = wsTarget3.Cells(i, 1).Value 
        found = False

        For j = 2 To lastRowSource
            If wsSource.Cells(j, 101).Value = codePart Then 
                With wsTarget3
                    .Cells(targetRow3, 1).Value = wsSource.Cells(j, 101).Value
                    .Cells(targetRow3, 2).Value = wsSource.Cells(j, 93).Value
                    .Cells(targetRow3, 3).Value = wsSource.Cells(j, 101).Value
                    .Cells(targetRow3, 4).Value = wsSource.Cells(j, 97).Value
                    .Cells(targetRow3, 5).Value = wsSource.Cells(j, 92).Value
                    .Cells(targetRow3, 6).Value = wsSource.Cells(j, 95).Value
                End With
                found = True
                targetRow3 = targetRow3 + 2
                Exit For
            End If
        Next j

        If Not found Then
            With wsTarget3
                .Cells(targetRow3, 1).Value = "Tidak Ditemukan"
                .Cells(targetRow3, 2).Resize(1, 6).ClearContents
            End With
            targetRow3 = targetRow3 + 2
        End If
    Next i

    Set targetRange = wsTarget3.Range("A3:F" & targetRow3 - 1)
    With targetRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

End Sub

Sub ColorRows()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim ws3 As Worksheet
    Dim lastRow1 As Long, lastRow2 As Long, lastRow3 As Long
    Dim i As Long

    Set ws1 = ThisWorkbook.Sheets("RESULT PARTLIST")
    Set ws2 = ThisWorkbook.Sheets("RESULT MARTLIST")
    Set ws3 = ThisWorkbook.Sheets("RESULT APL")

    lastRow1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    lastRow3 = ws3.Cells(ws3.Rows.Count, "A").End(xlUp).Row

    i = 2
    Do While i <= lastRow1 Or i <= lastRow2 Or i <= lastRow3
        If i + 1 <= lastRow1 Then
            ws1.Range(ws1.Cells(i, 1), ws1.Cells(i + 1, 1)).Interior.Color = RGB(211, 211, 211)
        End If
        If i + 1 <= lastRow2 Then
            ws2.Range(ws2.Cells(i, 1), ws2.Cells(i + 1, 1)).Interior.Color = RGB(211, 211, 211)
        End If
        If i + 1 <= lastRow3 Then
            ws3.Range(ws3.Cells(i, 1), ws3.Cells(i + 1, 1)).Interior.Color = RGB(211, 211, 211)
        End If

        If i + 2 <= lastRow1 Then
            ws1.Range(ws1.Cells(i + 2, 1), ws1.Cells(i + 3, 1)).Interior.Color = RGB(255, 255, 255)
        End If
        If i + 2 <= lastRow2 Then
            ws2.Range(ws2.Cells(i + 2, 1), ws2.Cells(i + 3, 1)).Interior.Color = RGB(255, 255, 255)
        End If
        If i + 2 <= lastRow3 Then
            ws3.Range(ws3.Cells(i + 2, 1), ws3.Cells(i + 3, 1)).Interior.Color = RGB(255, 255, 255)
        End If

        i = i + 4
    Loop
End Sub

Sub CompareAndHighlightDifferences()
    Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet
    Dim lastRow1 As Long, lastRow2 As Long, lastRow3 As Long
    Dim i As Long, col As Long
    Dim differences As String

    Set ws1 = ThisWorkbook.Sheets("RESULT PARTLIST")
    Set ws2 = ThisWorkbook.Sheets("RESULT MARTLIST")
    Set ws3 = ThisWorkbook.Sheets("RESULT APL")

    With ws1
        .Cells(1, 1).Value = "Code-P-SP-Part"
        .Cells(1, 2).Value = "Mat"
        .Cells(1, 3).Value = "ProfileType"
        .Cells(1, 4).Value = "Thk"
        .Cells(1, 5).Value = "Qty"
        .Cells(1, 6).Value = "Pos"
        .Cells(1, 7).Value = "Grade"
        .Cells(1, 8).Value = "Weight"
        .Cells(1, 9).Value = "NestingName"
        .Cells(1, 10).Value = "PS/PD"
        .Cells(1, 11).Value = "Stage"
        .Cells(1, 12).Value = "Note"
        .Cells(1, 14).Value = "Perbedaan"
        .Range("A1:L1").Font.Bold = True
        .Range("A1:L1").BorderAround ColorIndex:=1, Weight:=xlThin
        .Range("N1").Font.Bold = True
    End With

    With ws2
        .Cells(1, 1).Value = "NEST NAME"
        .Cells(1, 2).Value = "REMARK"
        .Cells(1, 3).Value = "QTY"
        .Cells(1, 4).Value = "LENGTH1"
        .Cells(1, 5).Value = "WIDTH1"
        .Cells(1, 6).Value = "THICK1"
        .Cells(1, 7).Value = "GRADE"
        .Cells(1, 8).Value = "NET"
        .Cells(1, 9).Value = "GROSS"
        .Cells(1, 10).Value = "REMNANT"
        .Cells(1, 11).Value = "THICK2"
        .Cells(1, 12).Value = "LENGTH2"
        .Cells(1, 13).Value = "WIDTH2"
        .Cells(1, 14).Value = "WEIGHT"
        .Cells(1, 16).Value = "PERBEDAAN"
        .Range("A1:N1").Font.Bold = True
        .Range("A1:N1").BorderAround ColorIndex:=1, Weight:=xlThin
        .Range("P1").Font.Bold = True
    End With

    With ws3
        .Cells(1, 1).Value = "Code-P-SP-Part"
        .Cells(1, 2).Value = "QTY"
        .Cells(1, 3).Value = "Drawing"
        .Cells(1, 4).Value = "Format"
        .Cells(1, 5).Value = "THK"
        .Cells(1, 6).Value = "Material"
        .Cells(1, 8).Value = "PERBEDAAN"
        .Range("A1:F1").Font.Bold = True
        .Range("A1:F1").BorderAround ColorIndex:=1, Weight:=xlThin
        .Range("H1").Font.Bold = True
    End With


    lastRow1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    lastRow3 = ws3.Cells(ws3.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow1 - 1 Step 2
        differences = ""
        For col = 4 To 9
            If ws1.Cells(i, col).Value <> ws1.Cells(i + 1, col).Value Then
                Select Case col
                    Case 4: differences = differences & "Thk berubah; "
                    Case 5: differences = differences & "Qty berubah; "
                    Case 6: differences = differences & "Pos berubah; "
                    Case 7: differences = differences & "Grade berubah; "
                    Case 8: differences = differences & "Weight berubah; "
                    Case 9: differences = differences & "NestingName berubah; "
                End Select
                ws1.Cells(i, col).Interior.Color = RGB(255, 255, 0)
                ws1.Cells(i + 1, col).Interior.Color = RGB(255, 255, 0)
            Else
                ws1.Cells(i, col).Interior.ColorIndex = xlNone
                ws1.Cells(i + 1, col).Interior.ColorIndex = xlNone
            End If
        Next col
        If differences <> "" Then
            ws1.Cells(i, 14).Value = Left(differences, Len(differences) - 2)
        Else
            ws1.Cells(i, 14).Value = ""
        End If
    Next i

    For i = 2 To lastRow2 - 1 Step 2
        differences = ""
        For col = 3 To 14
            If IsEmpty(ws2.Cells(i, col).Value) And IsEmpty(ws2.Cells(i + 1, col).Value) Then
                differences = differences & "REMNANT kosong; "
                differences = differences & "THICK2 kosong; "
                differences = differences & "LENGTH2 kosong; "
                differences = differences & "WIDTH2 kosong; "
                differences = differences & "WEIGHT kosong; "
                Exit For
            End If
            
            If ws2.Cells(i, col).Value <> ws2.Cells(i + 1, col).Value Then
                Select Case col
                    Case 3: differences = differences & "QTY berubah; "
                    Case 4: differences = differences & "LENGTH1 berubah; "
                    Case 5: differences = differences & "WIDTH1 berubah; "
                    Case 6: differences = differences & "THICK1 berubah; "
                    Case 7: differences = differences & "GRADE berubah; "
                    Case 8: differences = differences & "NET berubah; "
                    Case 9: differences = differences & "GROSS berubah;"
                    Case 10: differences = differences & "REMNANT berubah;"
                    Case 11: differences = differences & "THICK2 berubah; "
                    Case 12: differences = differences & "LENGTH2 berubah; "
                    Case 13: differences = differences & "WIDTH2 berubah; "
                    Case 14: differences = differences & "WEIGHT berubah; "
                End Select
                ws2.Cells(i, col).Interior.Color = RGB(255, 255, 0)
                ws2.Cells(i + 1, col).Interior.Color = RGB(255, 255, 0)
            Else
                ws2.Cells(i, col).Interior.ColorIndex = xlNone
                ws2.Cells(i + 1, col).Interior.ColorIndex = xlNone
            End If
        Next col
        If differences <> "" Then
            ws2.Cells(i, 16).Value = Left(differences, Len(differences) - 2)
        Else
            ws2.Cells(i, 16).Value = ""
        End If
    Next i

    For i = 2 To lastRow3 - 1 Step 2
        differences = ""
        For col = 2 To 6
            If ws3.Cells(i, col).Value <> ws3.Cells(i + 1, col).Value Then
                Select Case col
                    Case 2: differences = differences & "QTY berubah; "
                    Case 3: differences = differences & "Drawing berubah; "
                    Case 4: differences = differences & "Format berubah; "
                    Case 5: differences = differences & "THK berubah; "
                    Case 6: differences = differences & "Material berubah; "
                End Select
                ws3.Cells(i, col).Interior.Color = RGB(255, 255, 0)
                ws3.Cells(i + 1, col).Interior.Color = RGB(255, 255, 0)
            Else
                ws3.Cells(i, col).Interior.ColorIndex = xlNone
                ws3.Cells(i + 1, col).Interior.ColorIndex = xlNone
            End If
        Next col
        If differences <> "" Then
            ws3.Cells(i, 8).Value = Left(differences, Len(differences) - 2)
        Else
            ws3.Cells(i, 8).Value = ""
        End If
    Next i

End Sub

Sub AddPart()
    On Error GoTo ErrorHandler
    Dim wsSource As Worksheet
    Dim wsTarget1 As Worksheet
    Dim wsTarget2 As Worksheet
    Dim wsTarget3 As Worksheet
    Dim lastRow As Long
    Dim targetRow1 As Long
    Dim targetRow2 As Long
    Dim targetRow3 As Long
    Dim i As Long
    Dim targetRange As Range

    Set wsSource = ThisWorkbook.Sheets("ALL INPUT") 
    Set wsTarget1 = ThisWorkbook.Sheets("RESULT MARTLIST") 
    Set wsTarget2 = ThisWorkbook.Sheets("RESULT PARTLIST") 
    Set wsTarget3 = ThisWorkbook.Sheets("RESULT APL") 

    lastRow = wsSource.Cells(wsSource.Rows.Count, "AO").End(xlUp).Row
    targetRow1 = wsTarget1.Cells(wsTarget1.Rows.Count, "A").End(xlUp).Row + 1

    For i = 2 To lastRow
        If wsSource.Cells(i, 72).Value = "Unik" And wsSource.Cells(i, 41).Value <> "" Then
            wsSource.Range(wsSource.Cells(i, 41), wsSource.Cells(i, 54)).Copy Destination:=wsTarget1.Cells(targetRow1, 1)
            wsTarget1.Cells(targetRow1, 16).Value = "Add Part"
            wsTarget1.Range(wsTarget1.Cells(targetRow1, 1), wsTarget1.Cells(targetRow1, 14)).Interior.Color = RGB(0, 255, 0)
            targetRow1 = targetRow1 + 1
        End If
    Next i

    Set targetRange = wsTarget1.Range("A3:N" & targetRow1 - 1)
    With targetRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 0
    End With

    lastRow = wsSource.Cells(wsSource.Rows.Count, "CW").End(xlUp).Row
    targetRow3 = wsTarget3.Cells(wsTarget3.Rows.Count, "A").End(xlUp).Row + 1

    For i = 2 To lastRow
        If InStr(1, wsSource.Cells(i, 89).Value, "Code-P-SP-Part", vbTextCompare) = 0 Then
            If wsSource.Cells(i, 117).Value = "Unik" And wsSource.Cells(i, 92).Value <> "" Then
                wsTarget3.Cells(targetRow3, 1).Value = wsSource.Cells(i, 101).Value 
                wsTarget3.Cells(targetRow3, 2).Value = wsSource.Cells(i, 93).Value  
                wsTarget3.Cells(targetRow3, 3).Value = wsSource.Cells(i, 101).Value 
                wsTarget3.Cells(targetRow3, 4).Value = wsSource.Cells(i, 97).Value  
                wsTarget3.Cells(targetRow3, 5).Value = wsSource.Cells(i, 92).Value  
                wsTarget3.Cells(targetRow3, 6).Value = wsSource.Cells(i, 95).Value  
    
                wsTarget3.Cells(targetRow3, 8).Value = "Not Nesting"
                wsTarget3.Range(wsTarget3.Cells(targetRow3, 1), wsTarget3.Cells(targetRow3, 6)).Interior.Color = RGB(0, 255, 0)
    
                targetRow3 = targetRow3 + 1
            End If
        End If
    Next i

    Set targetRange = wsTarget3.Range("A3:F" & targetRow3 - 1)
    With targetRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 0
    End With
    
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    targetRow2 = wsTarget2.Cells(wsTarget2.Rows.Count, "A").End(xlUp).Row + 1

    For i = 2 To lastRow
        If wsSource.Cells(i, 27).Value = "Unik" And wsSource.Cells(i, 6).Value <> "" Then
            wsSource.Range(wsSource.Cells(i, 1), wsSource.Cells(i, 12)).Copy Destination:=wsTarget2.Cells(targetRow2, 1)
            wsTarget2.Cells(targetRow2, 14).Value = "Add Part"
            wsTarget2.Range(wsTarget2.Cells(targetRow2, 1), wsTarget2.Cells(targetRow2, 12)).Interior.Color = RGB(0, 255, 0)
            targetRow2 = targetRow2 + 1
        End If
    Next i

    Set targetRange = wsTarget2.Range("A3:L" & targetRow2 - 1)
    With targetRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 0
    End With


    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

Sub DeletePart()
    On Error GoTo ErrorHandler
    Dim wsSource As Worksheet
    Dim wsTarget1 As Worksheet
    Dim wsTarget2 As Worksheet
    Dim wsTarget3 As Worksheet
    Dim lastRow As Long
    Dim targetRow1 As Long
    Dim targetRow2 As Long
    Dim targetRow3 As Long
    Dim i As Long
    Dim targetRange As Range

    Set wsSource = ThisWorkbook.Sheets("ALL INPUT") 
    Set wsTarget1 = ThisWorkbook.Sheets("RESULT PARTLIST") 
    Set wsTarget2 = ThisWorkbook.Sheets("RESULT MARTLIST") 
    Set wsTarget3 = ThisWorkbook.Sheets("RESULT APL") 

    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

    targetRow1 = wsTarget1.Cells(wsTarget1.Rows.Count, "A").End(xlUp).Row + 1
    targetRow2 = wsTarget2.Cells(wsTarget2.Rows.Count, "A").End(xlUp).Row + 1
    targetRow3 = wsTarget3.Cells(wsTarget3.Rows.Count, "A").End(xlUp).Row + 1

    For i = 2 To lastRow 
        
        If wsSource.Cells(i, 28).Value = "Unik" And wsSource.Cells(i, 14).Value <> "" And wsSource.Cells(i, 19).Value <> "" Then
            wsSource.Range(wsSource.Cells(i, 14), wsSource.Cells(i, 25)).Copy Destination:=wsTarget1.Cells(targetRow1, 1)

            wsTarget1.Cells(targetRow1, 14).Value = "Delete Part"

            wsTarget1.Range(wsTarget1.Cells(targetRow1, 1), wsTarget1.Cells(targetRow1, 12)).Interior.Color = RGB(255, 0, 0)

            targetRow1 = targetRow1 + 1
        End If

        If wsSource.Cells(i, 72).Value = "Unik" And wsSource.Cells(i, 56).Value <> "" And wsSource.Cells(i, 65).Value <> "" Then
            wsSource.Range(wsSource.Cells(i, 57), wsSource.Cells(i, 69)).Copy Destination:=wsTarget2.Cells(targetRow2, 1)

            wsTarget2.Cells(targetRow2, 16).Value = "Delete Part"

            wsTarget2.Range(wsTarget2.Cells(targetRow2, 1), wsTarget2.Cells(targetRow2, 14)).Interior.Color = RGB(255, 0, 0)

            targetRow2 = targetRow2 + 1
        End If

        If wsSource.Cells(i, 118).Value = "Unik" And wsSource.Cells(i, 104).Value <> "" And wsSource.Cells(i, 109).Value <> "" Then
            wsSource.Range(wsSource.Cells(i, 104), wsSource.Cells(i, 109)).Copy Destination:=wsTarget3.Cells(targetRow3, 1)

            wsTarget3.Cells(targetRow3, 8).Value = "Delete Part"

            wsTarget3.Range(wsTarget3.Cells(targetRow3, 1), wsTarget3.Cells(targetRow3, 6)).Interior.Color = RGB(255, 0, 0)

            targetRow3 = targetRow3 + 1
        End If
    Next i

    If targetRow1 > 3 Then
        Set targetRange = wsTarget1.Range("A3:L" & targetRow1 - 1)
        With targetRange.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = 0
        End With
    End If

    If targetRow2 > 3 Then
        Set targetRange = wsTarget2.Range("A3:N" & targetRow2 - 1)
        With targetRange.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = 0
        End With
    End If

    If targetRow3 > 3 Then
        Set targetRange = wsTarget3.Range("A3:F" & targetRow3 - 1)
        With targetRange.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = 0
        End With
    End If

    MsgBox "Proses selesai!", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub
Sub DeleteRowsAndCleanColumnP()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim originalText As String
    Dim newText As String
    Dim targetString As String
    Dim keywords As Variant
    Dim keyword As Variant
    Dim containsKeyword As Boolean
    
    Set ws = ThisWorkbook.Sheets("RESULT MARTLIST") 
    lastRow = ws.Cells(ws.Rows.Count, "P").End(xlUp).Row
    
    targetString = "REMNANT kosong; THICK2 kosong; LENGTH2 kosong; WIDTH2 kosong; WEIGHT kosong"
    
    keywords = Array("QTY berubah; ", "LENGTH1 berubah; ", "WIDTH1 berubah; ", "THICK1 berubah; ", _
                     "GRADE berubah; ", "NET berubah; ", "GROSS berubah; ", "REMNANT berubah; ", _
                     "THICK2 berubah; ", "LENGTH2 berubah; ", "WIDTH2 berubah; ", "WEIGHT berubah; ")
    
    For i = lastRow To 1 Step -1 
        If ws.Cells(i, 16).Value = targetString Then 
            ws.Rows(i & ":" & i + 1).Delete
        End If
    Next i
    
    lastRow = ws.Cells(ws.Rows.Count, "P").End(xlUp).Row 
    For i = 1 To lastRow
        originalText = ws.Cells(i, 16).Value 
        
        containsKeyword = False
        For Each keyword In keywords
            If InStr(originalText, keyword) > 0 Then
                containsKeyword = True
                Exit For
            End If
        Next keyword
        
        newText = Replace(originalText, "REMNANT kosong; ", "")
        newText = Replace(newText, "THICK2 kosong; ", "")
        newText = Replace(newText, "LENGTH2 kosong; ", "")
        newText = Replace(newText, "WIDTH2 kosong; ", "")
        newText = Replace(newText, "WEIGHT kosong", "")
        
        ws.Cells(i, 16).Value = Trim(newText)
    Next i
End Sub

