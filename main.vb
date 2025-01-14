Sub RUN_COMPARE()
    Call RoundInPlace
    Call RemoveSpacesAllCells
    Call Code_P_SP_Part_Pos
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

    ' Set worksheet yang aktif
    Set ws = ActiveSheet

    ' Definisikan rentang kolom yang akan diproses
    colRanges = Array("A:Y", "AN:BR", "CK:DJ")

    Application.ScreenUpdating = False

    ' Loop untuk setiap rentang kolom yang telah ditentukan
    For Each colRange In colRanges
        ' Pisahkan rentang menjadi kolom awal dan kolom akhir
        startCol = Split(colRange, ":")(0)
        endCol = Split(colRange, ":")(1)

        ' Loop untuk setiap kolom dalam rentang
        Dim col As Range
        For Each col In ws.Range(startCol & "1:" & endCol & ws.Rows.Count).Columns
            ' Cari baris terakhir pada kolom yang sedang diproses
            lastRow = ws.Cells(ws.Rows.Count, col.Column).End(xlUp).Row

            ' Tetapkan range berdasarkan kolom dan baris terakhir
            Set rng = ws.Range(ws.Cells(1, col.Column), ws.Cells(lastRow, col.Column))

            ' Loop untuk setiap sel dalam rentang kolom
            For Each cell In rng
                If Not IsEmpty(cell.Value) Then
                    ' Jika sel berupa angka, biarkan tetap angka
                    If IsNumeric(cell.Value) Then
                        cell.Value = cell.Value * 1 ' Pastikan tetap berupa angka
                    Else
                        cell.Value = Trim(cell.Value) ' Hapus spasi jika bukan angka
                    End If
                End If
            Next cell
        Next col
    Next colRange

    Application.ScreenUpdating = True

    'MsgBox "Proses penghapusan spasi selesai!", vbInformation
End Sub

Sub Code_P_SP_Part_Pos()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Mengatur worksheet "ALL INPUT"
    Set ws = ThisWorkbook.Sheets("ALL INPUT")
    
    ' Menentukan baris terakhir di kolom CK
    lastRow = ws.Cells(ws.Rows.Count, "CK").End(xlUp).Row
    
    ' Menggabungkan data dari kolom CK dan CP ke kolom CW
    For i = 6 To lastRow
        ws.Cells(i, "CW").Value = ws.Cells(i, "CK").Value & "-" & ws.Cells(i, "CP").Value
    Next i
    
    ' MsgBox "Penggabungan data selesai!", vbInformation
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

    ' ================= Program 1: Copy ke RESULT PARTLIST =================
    ' Set worksheets
    Set wsSource = ThisWorkbook.Sheets("ALL INPUT") ' Sheet sumber
    Set wsTargetPartlist = ThisWorkbook.Sheets("RESULT PARTLIST") ' Sheet tujuan PARTLIST
    
    ' Cari baris terakhir di Sheet sumber
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    
    ' Awal baris untuk menulis di Sheet tujuan PARTLIST
    targetRowPartlist = 2
    
    ' Kolom untuk memeriksa "Unik" (AC hingga AH = kolom 29 sampai 34)
    colStart = 29
    colEnd = 34
    
    ' Daftar teks yang harus diabaikan
    skipText = Array("Block", "Block Code", "Panel", "Code-P-SP-Part")

    ' Looping melalui setiap baris di Sheet sumber
    For i = 2 To lastRow
        uniqueFound = False
        matchFound = False
        
        ' Periksa jika kolom AB berisi "Duplikat"
        If wsSource.Cells(i, 28).Value = "Duplikat" Then ' Kolom AB = kolom ke-28
            ' Cek jika salah satu kolom AC hingga AH bernilai "Unik"
            For j = colStart To colEnd
                If wsSource.Cells(i, j).Value = "Unik" Then
                    uniqueFound = True
                    Exit For
                End If
            Next j
            
            ' Jika "Unik" ditemukan, cek kecocokan antara kolom N dan kolom A
            If uniqueFound Then
                codePart = wsSource.Cells(i, 14).Value ' Kolom N = kolom ke-14
                
                ' Periksa apakah kolom N berisi salah satu teks yang harus diabaikan
                Dim textFound As Boolean
                textFound = False
                For Each Skip In skipText
                    If InStr(1, codePart, Skip, vbTextCompare) > 0 Then
                        textFound = True
                        Exit For
                    End If
                Next Skip

                ' Jika teks ditemukan, lewati baris ini
                If textFound Then GoTo NextRowPartlist
                
                ' Loop untuk mencari kecocokan di kolom A
                For j = 2 To lastRow
                    If wsSource.Cells(j, 1).Value = codePart Then ' Kolom A = kolom ke-1
                        matchFound = True
                        Exit For
                    End If
                Next j
                
                ' Jika kecocokan ditemukan, salin kolom A hingga AH ke RESULT PARTLIST
                If matchFound Then
                    wsTargetPartlist.Range(wsTargetPartlist.Cells(targetRowPartlist, 1), wsTargetPartlist.Cells(targetRowPartlist, 34)).Value = _
                        wsSource.Range(wsSource.Cells(i, 1), wsSource.Cells(i, 34)).Value
                    targetRowPartlist = targetRowPartlist + 1
                End If
            End If
        End If
NextRowPartlist:
    Next i
    
    ' ================= Program 2: Copy ke RESULT MARTLIST =================
    ' Set worksheet tujuan untuk MARTLIST
    Set wsTargetMartlist = ThisWorkbook.Sheets("RESULT MARTLIST")
    
    ' Awal baris untuk menulis di Sheet tujuan MARTLIST
    targetRowMartlist = 2
    
    ' Kolom untuk memeriksa "Unik" (BV hingga CG = kolom 74 sampai 85)
    colStart = 74
    colEnd = 85

    ' Looping melalui setiap baris di Sheet sumber
    For i = 2 To lastRow
        uniqueFound = False
        matchFound = False
        
        ' Periksa jika kolom BT berisi "Duplikat"
        If wsSource.Cells(i, 73).Value = "Duplikat" Then ' Kolom BU = kolom ke-73
            ' Cek jika salah satu kolom BU hingga CF bernilai "Unik"
            For j = colStart To colEnd
                If wsSource.Cells(i, j).Value = "Unik" Then
                    uniqueFound = True
                    Exit For
                End If
            Next j
            
            ' Jika "Unik" ditemukan, cek kecocokan antara kolom BE dan kolom AO
            If uniqueFound Then
                codePart = wsSource.Cells(i, 57).Value ' Kolom BE = kolom ke-57
                
                ' Loop untuk mencari kecocokan di kolom AN
                For j = 2 To lastRow
                    If wsSource.Cells(j, 41).Value = codePart Then ' Kolom AN = kolom ke-40
                        matchFound = True
                        Exit For
                    End If
                Next j
                
                ' Jika kecocokan ditemukan, salin kolom AN hingga CG ke RESULT MARTLIST
                If matchFound Then
                    wsTargetMartlist.Range(wsTargetMartlist.Cells(targetRowMartlist, 1), wsTargetMartlist.Cells(targetRowMartlist, 46)).Value = _
                        wsSource.Range(wsSource.Cells(i, 40), wsSource.Cells(i, 85)).Value
                    targetRowMartlist = targetRowMartlist + 1
                End If
            End If
        End If
    Next i

        ' ================= Program 3: Copy ke RESULT APL =================
    ' Set worksheet tujuan untuk APL
    Set wsTargetMartlist = ThisWorkbook.Sheets("RESULT APL")
    
    ' Awal baris untuk menulis di Sheet tujuan APL
    targetRowMartlist = 2
    
    ' Kolom untuk memeriksa "Unik" (DO hingga DR = kolom 119 sampai 122)
    colStart = 119
    colEnd = 122

    ' Looping melalui setiap baris di Sheet sumber
    For i = 2 To lastRow
        uniqueFound = False
        matchFound = False
        
        ' Periksa jika kolom BT berisi "Duplikat"
        If wsSource.Cells(i, 118).Value = "Duplikat" Then ' Kolom DN = kolom ke-118
            ' Cek jika salah satu kolom DO hingga DR bernilai "Unik"
            For j = colStart To colEnd
                If wsSource.Cells(i, j).Value = "Unik" Then
                    uniqueFound = True
                    Exit For
                End If
            Next j
            
            ' Jika "Unik" ditemukan, cek kecocokan antara kolom CZ dan kolom CW
            If uniqueFound Then
                codePart = wsSource.Cells(i, 104).Value ' Kolom CZ = kolom ke-104
                
                ' Loop untuk mencari kecocokan di kolom CZ
                For j = 2 To lastRow
                    If wsSource.Cells(j, 101).Value = codePart Then ' Kolom CW = kolom ke-101
                        matchFound = True
                        Exit For
                    End If
                Next j
                
                ' Jika kecocokan ditemukan, salin kolom CK hingga DR ke RESULT MARTLIST
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
    
    ' Set worksheet untuk RESULT PARTLIST
    Set ws1 = ThisWorkbook.Sheets("RESULT PARTLIST")
    ' Set worksheet untuk RESULT MARTLIST
    Set ws2 = ThisWorkbook.Sheets("RESULT MARTLIST")
    ' Set worksheet untuk RESULT APL
    Set ws3 = ThisWorkbook.Sheets("RESULT APL")
    
    ' Menghapus semua isi, format, dan komentar dari kolom A hingga L dan AA hingga DS di RESULT PARTLIST
    ws1.Columns("A:L").Clear
    ws1.Columns("AA:AH").Clear
    
    ' Menghapus semua isi, format, dan komentar dari kolom A hingga Q dan AG hingga AT di RESULT MARTLIST
    ws2.Columns("A:Q").Clear
    ws2.Columns("AG:AT").Clear

    ' Menghapus semua isi, format, dan komentar dari kolom A hingga M dan AC hingga AT di RESULT APL
    ws3.Columns("A:M").Clear
    ws3.Columns("X:AT").Clear
    
    ' Cek kolom T (kolom 20) di RESULT MARTLIST dan hapus data di kolom R hingga AT jika kolom T kosong
    ' Cari baris terakhir dengan data di sheet RESULT MARTLIST
    lastRow = ws2.Cells(ws2.Rows.Count, "T").End(xlUp).Row
    
    ' Looping dari baris 2 (asumsi baris 1 adalah header) hingga baris terakhir
    For i = 2 To lastRow
        If ws2.Cells(i, 20).Value = "" Then ' Periksa jika kolom T (20) kosong
            ' Hapus data atau format di kolom R (18) hingga AT (31) pada baris tersebut
            ws2.Range(ws2.Cells(i, 18), ws2.Cells(i, 31)).Clear
        End If
    Next i

    ' Cari barisan terakhir pada semua data di RESULT MARTLIST
    lastRow = ws2.Cells.Find(What:="*", LookIn:=xlFormulas, _
                              SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

    ' Tentukan barisan akhir untuk dihapus
    If lastRow >= 2 Then
        rowToStart = lastRow - 1 ' Mulai dari dua barisan terakhir
        ' Hapus isi data pada dua barisan terakhir
        ws2.Rows(rowToStart & ":" & lastRow).ClearContents
        ' MsgBox "Isi data pada 2 baris terakhir telah dihapus.", vbInformation
    ElseIf lastRow = 1 Then
        ' Jika hanya ada satu barisan data
        ws2.Rows(lastRow).ClearContents
        ' MsgBox "Hanya satu baris yang tersedia, isi data telah dihapus.", vbInformation
    Else
        ' MsgBox "Tidak ada data yang bisa dihapus.", vbExclamation
    End If
End Sub

Sub MoveDataToEvenRowsAndClearSource()
    Dim wsTarget1 As Worksheet
    Dim wsTarget2 As Worksheet
    Dim wsTarget3 As Worksheet
    Dim lastRow1 As Long, lastRow2 As Long, lastRow3 As Long
    Dim targetRow1 As Long, targetRow2 As Long, targetRow3 As Long
    Dim i As Long

    ' Set worksheet tujuan
    Set wsTarget1 = ThisWorkbook.Sheets("RESULT PARTLIST") ' Sheet tujuan 1
    Set wsTarget2 = ThisWorkbook.Sheets("RESULT MARTLIST")     ' Sheet tujuan 2
    Set wsTarget3 = ThisWorkbook.Sheets("RESULT APL")     ' Sheet tujuan 3

    ' 1. Pindahkan data dari kolom N:Y ke RESULT UNIK PARTLIST
    ' Cari baris terakhir di kolom N
    lastRow1 = wsTarget1.Cells(wsTarget1.Rows.Count, "N").End(xlUp).Row ' Kolom N
    
    ' Awal baris genap untuk RESULT UNIK PARTLIST
    targetRow1 = 2 ' Baris genap pertama
    
    For i = 2 To lastRow1
        If wsTarget1.Cells(i, 14).Value <> "" Then ' Kolom N = 14
            ' Salin kolom N sampai Y ke kolom A sampai L
            wsTarget1.Range(wsTarget1.Cells(i, 14), wsTarget1.Cells(i, 25)).Copy _
                Destination:=wsTarget1.Cells(targetRow1, 1)
            
            ' Tingkatkan targetRow1 ke baris genap berikutnya
            targetRow1 = targetRow1 + 2
        End If
    Next i
    
    ' Hapus data sumber di kolom N sampai Y
    wsTarget1.Range("N2:Y" & lastRow1).Clear
    
    ' 2. Pindahkan data dari kolom R:AE ke RESULT MARTLIST
    ' Cari baris terakhir di kolom R
    lastRow2 = wsTarget2.Cells(wsTarget2.Rows.Count, "R").End(xlUp).Row ' Kolom R
    
    ' Awal baris genap untuk RESULT MARTLIST
    targetRow2 = 2 ' Baris genap pertama
    
    For i = 2 To lastRow2
        If wsTarget2.Cells(i, 18).Value <> "" Then ' Kolom R = 18
            ' Salin kolom R sampai AE ke kolom A sampai N
            wsTarget2.Range(wsTarget2.Cells(i, 18), wsTarget2.Cells(i, 31)).Copy _
                Destination:=wsTarget2.Cells(targetRow2, 1)
            
            ' Tingkatkan targetRow2 ke baris genap berikutnya
            targetRow2 = targetRow2 + 2
        End If
    Next i
    
    ' Hapus data sumber di kolom R sampai AE
    wsTarget2.Range("R2:AE" & lastRow2).Clear

    ' 3. Pindahkan data dari kolom P:X ke RESULT APL
    ' Cari baris terakhir di kolom P
    lastRow3 = wsTarget3.Cells(wsTarget3.Rows.Count, "P").End(xlUp).Row ' Kolom P
    
    ' Awal baris genap untuk RESULT APL
    targetRow3 = 2 ' Baris genap pertama
    
    For i = 2 To lastRow3
        If wsTarget3.Cells(i, 16).Value <> "" Then ' Kolom P = 16
            ' Salin kolom P sampai X ke kolom A sampai I
            wsTarget3.Range(wsTarget3.Cells(i, 16), wsTarget3.Cells(i, 21)).Copy _
                Destination:=wsTarget3.Cells(targetRow3, 1)
            
            ' Tingkatkan targetRow3 ke baris genap berikutnya
            targetRow3 = targetRow3 + 2
        End If
    Next i
    
    ' Hapus data sumber di kolom P sampai X
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

    ' ======== Sheet Setup ========
    Set wsSource = ThisWorkbook.Sheets("ALL INPUT") ' Sheet sumber
    Set wsTarget1 = ThisWorkbook.Sheets("RESULT PARTLIST") ' Sheet tujuan pertama
    Set wsTarget2 = ThisWorkbook.Sheets("RESULT MARTLIST") ' Sheet tujuan kedua
    Set wsTarget3 = ThisWorkbook.Sheets("RESULT APL") ' Sheet tujuan ketiga

    ' ======== Baris Terakhir ========
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row ' Sheet sumber
    lastRowTarget1 = wsTarget1.Cells(wsTarget1.Rows.Count, "A").End(xlUp).Row ' Sheet tujuan pertama
    lastRowTarget2 = wsTarget2.Cells(wsTarget2.Rows.Count, "A").End(xlUp).Row ' Sheet tujuan kedua
    lastRowTarget3 = wsTarget3.Cells(wsTarget3.Rows.Count, "A").End(xlUp).Row ' Sheet tujuan ketiga

    ' ======== Proses untuk RESULT PARTLIST ========
    targetRow1 = 3 ' Mulai dari baris ganjil pertama di RESULT PARTLIST
    For i = 2 To lastRowTarget1 Step 2 ' Proses baris genap
        codePart = wsTarget1.Cells(i, 1).Value ' Ambil kode part dari kolom A
        found = False

        ' Loop pencarian di sheet sumber
        For j = 2 To lastRowSource
            If wsSource.Cells(j, 1).Value = codePart Then ' Kolom A = kolom ke-1
                With wsTarget1
                    .Cells(targetRow1, 1).Value = wsSource.Cells(j, 1).Value
                    .Cells(targetRow1, 2).Resize(1, 11).Value = wsSource.Cells(j, 2).Resize(1, 11).Value
                End With
                found = True
                targetRow1 = targetRow1 + 2
                Exit For
            End If
        Next j

        ' Jika tidak ditemukan
        If Not found Then
            With wsTarget1
                .Cells(targetRow1, 1).Value = "Tidak Ditemukan"
                .Cells(targetRow1, 2).Resize(1, 11).ClearContents
            End With
            targetRow1 = targetRow1 + 2
        End If
    Next i

    ' Tambahkan garis batas untuk RESULT PARTLIST
    Set targetRange = wsTarget1.Range("A3:L" & targetRow1 - 1)
    With targetRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    ' ======== Proses untuk RESULT MARTLIST ========
    targetRow2 = 3 ' Mulai dari baris ganjil pertama di RESULT MARTLIST
    For i = 2 To lastRowTarget2 Step 2 ' Proses baris genap
        codePart = wsTarget2.Cells(i, 1).Value ' Ambil kode part dari kolom A
        found = False

        ' Loop pencarian di sheet sumber
        For j = 2 To lastRowSource
            If wsSource.Cells(j, 41).Value = codePart Then ' Kolom AO = kolom ke-41
                With wsTarget2
                    .Cells(targetRow2, 1).Value = wsSource.Cells(j, 41).Value
                    .Cells(targetRow2, 2).Resize(1, 14).Value = wsSource.Cells(j, 42).Resize(1, 14).Value
                End With
                found = True
                targetRow2 = targetRow2 + 2
                Exit For
            End If
        Next j

        ' Jika tidak ditemukan
        If Not found Then
            With wsTarget2
                .Cells(targetRow2, 1).Value = "Tidak Ditemukan"
                .Cells(targetRow2, 2).Resize(1, 14).ClearContents
            End With
            targetRow2 = targetRow2 + 2
        End If
    Next i

    ' Tambahkan garis batas untuk RESULT MARTLIST
    Set targetRange = wsTarget2.Range("A3:N" & targetRow2 - 1)
    With targetRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    ' ======== Proses untuk RESULT APL ========
    targetRow3 = 3 ' Mulai dari baris ganjil pertama di RESULT APL
    For i = 2 To lastRowTarget3 Step 2 ' Proses baris genap
        codePart = wsTarget3.Cells(i, 1).Value ' Ambil kode part dari kolom A
        found = False

        ' Loop pencarian di sheet sumber
        For j = 2 To lastRowSource
            If wsSource.Cells(j, 101).Value = codePart Then ' Kolom CK = kolom ke-101
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

        ' Jika tidak ditemukan
        If Not found Then
            With wsTarget3
                .Cells(targetRow3, 1).Value = "Tidak Ditemukan"
                .Cells(targetRow3, 2).Resize(1, 6).ClearContents
            End With
            targetRow3 = targetRow3 + 2
        End If
    Next i

    ' Tambahkan garis batas untuk RESULT APL
    Set targetRange = wsTarget3.Range("A3:F" & targetRow3 - 1)
    With targetRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    'MsgBox "Proses selesai!", vbInformation
End Sub

Sub ColorRows()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim ws3 As Worksheet
    Dim lastRow1 As Long, lastRow2 As Long, lastRow3 As Long
    Dim i As Long

    ' Set worksheet
    Set ws1 = ThisWorkbook.Sheets("RESULT PARTLIST")
    Set ws2 = ThisWorkbook.Sheets("RESULT MARTLIST")
    Set ws3 = ThisWorkbook.Sheets("RESULT APL")

    ' Cari baris terakhir di kedua sheet
    lastRow1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    lastRow3 = ws3.Cells(ws3.Rows.Count, "A").End(xlUp).Row

    ' Looping melalui setiap empat barisan untuk mewarnai di kedua sheet
    i = 2
    Do While i <= lastRow1 Or i <= lastRow2 Or i <= lastRow3
        ' Warna abu-abu muda untuk dua barisan pertama
        If i + 1 <= lastRow1 Then
            ws1.Range(ws1.Cells(i, 1), ws1.Cells(i + 1, 1)).Interior.Color = RGB(211, 211, 211)
        End If
        If i + 1 <= lastRow2 Then
            ws2.Range(ws2.Cells(i, 1), ws2.Cells(i + 1, 1)).Interior.Color = RGB(211, 211, 211)
        End If
        If i + 1 <= lastRow3 Then
            ws3.Range(ws3.Cells(i, 1), ws3.Cells(i + 1, 1)).Interior.Color = RGB(211, 211, 211)
        End If

        ' Warna putih untuk dua barisan berikutnya
        If i + 2 <= lastRow1 Then
            ws1.Range(ws1.Cells(i + 2, 1), ws1.Cells(i + 3, 1)).Interior.Color = RGB(255, 255, 255)
        End If
        If i + 2 <= lastRow2 Then
            ws2.Range(ws2.Cells(i + 2, 1), ws2.Cells(i + 3, 1)).Interior.Color = RGB(255, 255, 255)
        End If
        If i + 2 <= lastRow3 Then
            ws3.Range(ws3.Cells(i + 2, 1), ws3.Cells(i + 3, 1)).Interior.Color = RGB(255, 255, 255)
        End If

        ' Geser ke bawah 4 langkah untuk memproses kelompok berikutnya
        i = i + 4
    Loop
End Sub

Sub CompareAndHighlightDifferences()
    Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet
    Dim lastRow1 As Long, lastRow2 As Long, lastRow3 As Long
    Dim i As Long, col As Long
    Dim differences As String

    ' Set worksheet untuk PARTLIST dan MARTLIST
    Set ws1 = ThisWorkbook.Sheets("RESULT PARTLIST")
    Set ws2 = ThisWorkbook.Sheets("RESULT MARTLIST")
    Set ws3 = ThisWorkbook.Sheets("RESULT APL")

    ' Tambahkan header untuk PARTLIST
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

    ' Tambahkan header untuk MARTLIST
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

     ' Tambahkan header untuk APL
    With ws3
        .Cells(1, 1).Value = "Code-P-SP-Part"
        .Cells(1, 2).Value = "QTY"
        .Cells(1, 3).Value = "Drawing"
        .Cells(1, 4).Value = "Format"
        .Cells(1, 5).Value = "THK"
        .Cells(1, 6).Value = "Material"
        .Cells(1, 8).Value = "Perbedaan"
        .Range("A1:F1").Font.Bold = True
        .Range("A1:F1").BorderAround ColorIndex:=1, Weight:=xlThin
        .Range("H1").Font.Bold = True
    End With


    ' Cari barisan terakhir untuk masing-masing worksheet
    lastRow1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    lastRow3 = ws3.Cells(ws3.Rows.Count, "A").End(xlUp).Row

    ' Perbandingan untuk PARTLIST
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


    ' Perbandingan untuk MARTLIST
    For i = 2 To lastRow2 - 1 Step 2
        differences = ""
        For col = 3 To 14
            ' Logika untuk mengecek jika tidak ada isi dalam MARTLIST
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

    ' Perbandingan untuk APL
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

    ' Set worksheets
    Set wsSource = ThisWorkbook.Sheets("ALL INPUT") ' Sheet sumber
    Set wsTarget1 = ThisWorkbook.Sheets("RESULT MARTLIST") ' Sheet tujuan untuk operasi pertama
    Set wsTarget2 = ThisWorkbook.Sheets("RESULT PARTLIST") ' Sheet tujuan untuk operasi kedua
    Set wsTarget3 = ThisWorkbook.Sheets("RESULT APL") ' Sheet tujuan untuk operasi kedua

    ' Bagian pertama: Operasi ke Sheet RESULT MARTLIST
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

    ' Bagian kedua: Operasi ke Sheet RESULT APL
    lastRow = wsSource.Cells(wsSource.Rows.Count, "CW").End(xlUp).Row
    targetRow3 = wsTarget3.Cells(wsTarget3.Rows.Count, "A").End(xlUp).Row + 1

    For i = 2 To lastRow
    ' Periksa jika kolom CK berisi teks "Code-P-SP-Part"
        If InStr(1, wsSource.Cells(i, 89).Value, "Code-P-SP-Part", vbTextCompare) = 0 Then
            If wsSource.Cells(i, 117).Value = "Unik" And wsSource.Cells(i, 92).Value <> "" Then
                ' Ambil data dari kolom tertentu di sheet sumber dan tempatkan di sheet tujuan
                wsTarget3.Cells(targetRow3, 1).Value = wsSource.Cells(i, 101).Value ' Kolom 101 ke Kolom A
                wsTarget3.Cells(targetRow3, 2).Value = wsSource.Cells(i, 93).Value  ' Kolom 93 ke Kolom B
                wsTarget3.Cells(targetRow3, 3).Value = wsSource.Cells(i, 101).Value ' Kolom 101 ke Kolom C
                wsTarget3.Cells(targetRow3, 4).Value = wsSource.Cells(i, 97).Value  ' Kolom 97 ke Kolom D
                wsTarget3.Cells(targetRow3, 5).Value = wsSource.Cells(i, 92).Value  ' Kolom 92 ke Kolom E
                wsTarget3.Cells(targetRow3, 6).Value = wsSource.Cells(i, 95).Value  ' Kolom 95 ke Kolom F
    
                ' Tandai baris dengan "Add Part" dan beri warna
                wsTarget3.Cells(targetRow3, 8).Value = "Not Nesting"
                wsTarget3.Range(wsTarget3.Cells(targetRow3, 1), wsTarget3.Cells(targetRow3, 6)).Interior.Color = RGB(0, 255, 0)
    
                targetRow3 = targetRow3 + 1
            End If
        End If
    Next i

    ' Tambahkan border pada rentang yang baru ditambahkan
    Set targetRange = wsTarget3.Range("A3:F" & targetRow3 - 1)
    With targetRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 0
    End With
    
    ' Bagian ketiga: Operasi ke Sheet RESULT PARTLIST
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

    ' Set worksheets
    Set wsSource = ThisWorkbook.Sheets("ALL INPUT") ' Sheet sumber
    Set wsTarget1 = ThisWorkbook.Sheets("RESULT PARTLIST") ' Sheet tujuan pertama
    Set wsTarget2 = ThisWorkbook.Sheets("RESULT MARTLIST") ' Sheet tujuan kedua
    Set wsTarget3 = ThisWorkbook.Sheets("RESULT APL") ' Sheet tujuan kedua

    ' Cari baris terakhir di Sheet sumber
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

    ' Awal baris untuk menulis di Sheet tujuan pertama
    targetRow1 = wsTarget1.Cells(wsTarget1.Rows.Count, "A").End(xlUp).Row + 1
    ' Awal baris untuk menulis di Sheet tujuan kedua
    targetRow2 = wsTarget2.Cells(wsTarget2.Rows.Count, "A").End(xlUp).Row + 1
    ' Awal baris untuk menulis di Sheet tujuan kedua
    targetRow3 = wsTarget3.Cells(wsTarget3.Rows.Count, "A").End(xlUp).Row + 1

    ' Looping melalui setiap baris di Sheet sumber
    For i = 2 To lastRow ' Mulai dari baris kedua (asumsi baris pertama adalah header)
        
        ' Kondisi pertama untuk memindahkan data ke RESULT PARTLIST
        If wsSource.Cells(i, 28).Value = "Unik" And wsSource.Cells(i, 14).Value <> "" And wsSource.Cells(i, 19).Value <> "" Then
            ' Salin kolom N sampai Y ke Sheet tujuan (kolom A sampai L)
            wsSource.Range(wsSource.Cells(i, 14), wsSource.Cells(i, 25)).Copy Destination:=wsTarget1.Cells(targetRow1, 1)

            ' Tambahkan keterangan "DEL PART" di kolom M tujuan
            wsTarget1.Cells(targetRow1, 14).Value = "Delete Part"

            ' Beri warna latar belakang light gold 2 di kolom A sampai L
            wsTarget1.Range(wsTarget1.Cells(targetRow1, 1), wsTarget1.Cells(targetRow1, 12)).Interior.Color = RGB(255, 0, 0)

            ' Tingkatkan targetRow untuk menulis di barisan berikutnya
            targetRow1 = targetRow1 + 1
        End If

        ' Kondisi kedua untuk memindahkan data ke RESULT MARTLIST
        If wsSource.Cells(i, 73).Value = "Unik" And wsSource.Cells(i, 56).Value <> "" And wsSource.Cells(i, 65).Value <> "" Then
            ' Salin kolom N sampai Y ke Sheet tujuan (kolom A sampai L)
            wsSource.Range(wsSource.Cells(i, 57), wsSource.Cells(i, 69)).Copy Destination:=wsTarget2.Cells(targetRow2, 1)

            ' Tambahkan keterangan "DEL PART" di kolom M tujuan
            wsTarget2.Cells(targetRow2, 16).Value = "Delete Part"

            ' Beri warna latar belakang light gold 2 di kolom A sampai L
            wsTarget2.Range(wsTarget2.Cells(targetRow2, 1), wsTarget2.Cells(targetRow2, 14)).Interior.Color = RGB(255, 0, 0)

            ' Tingkatkan targetRow untuk menulis di barisan berikutnya
            targetRow2 = targetRow2 + 1
        End If

        ' Kondisi ketiga untuk memindahkan data ke RESULT APL
        If wsSource.Cells(i, 118).Value = "Unik" And wsSource.Cells(i, 104).Value <> "" And wsSource.Cells(i, 109).Value <> "" Then
            ' Salin kolom N sampai Y ke Sheet tujuan (kolom A sampai L)
            wsSource.Range(wsSource.Cells(i, 104), wsSource.Cells(i, 109)).Copy Destination:=wsTarget3.Cells(targetRow3, 1)

            ' Tambahkan keterangan "DEL PART" di kolom M tujuan
            wsTarget3.Cells(targetRow3, 8).Value = "Delete Part"

            ' Beri warna latar belakang light gold 2 di kolom A sampai L
            wsTarget3.Range(wsTarget3.Cells(targetRow3, 1), wsTarget3.Cells(targetRow3, 6)).Interior.Color = RGB(255, 0, 0)

            ' Tingkatkan targetRow untuk menulis di barisan berikutnya
            targetRow3 = targetRow3 + 1
        End If
    Next i

    ' Tambahkan garis batas untuk kedua sheet jika ada data
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
    
    ' Set worksheet yang aktif
    Set ws = ThisWorkbook.Sheets("RESULT MARTLIST") ' Ganti dengan nama sheet Anda
    lastRow = ws.Cells(ws.Rows.Count, "P").End(xlUp).Row
    
    ' String yang dicari untuk menghapus baris
    targetString = "REMNANT kosong; THICK2 kosong; LENGTH2 kosong; WIDTH2 kosong; WEIGHT kosong"
    
    ' Array dari kata kunci yang tidak boleh dihapus
    keywords = Array("QTY berubah; ", "LENGTH1 berubah; ", "WIDTH1 berubah; ", "THICK1 berubah; ", _
                     "GRADE berubah; ", "NET berubah; ", "GROSS berubah; ", "REMNANT berubah; ", _
                     "THICK2 berubah; ", "LENGTH2 berubah; ", "WIDTH2 berubah; ", "WEIGHT berubah; ")
    
    ' Loop dari baris terakhir ke baris pertama untuk menghapus baris
    For i = lastRow To 1 Step -1 ' Mulai dari baris terakhir
        If ws.Cells(i, 16).Value = targetString Then ' Kolom P adalah kolom ke-16
            ' Hapus baris saat ini dan satu baris di bawahnya
            ws.Rows(i & ":" & i + 1).Delete
        End If
    Next i
    
    ' Loop melalui setiap baris di kolom P untuk menghapus bagian yang tidak diinginkan
    lastRow = ws.Cells(ws.Rows.Count, "P").End(xlUp).Row ' Update lastRow setelah penghapusan
    For i = 1 To lastRow
        originalText = ws.Cells(i, 16).Value ' Kolom P adalah kolom ke-16
        
        ' Cek apakah isi sel mengandung salah satu kata kunci
        containsKeyword = False
        For Each keyword In keywords
            If InStr(originalText, keyword) > 0 Then
                containsKeyword = True
                Exit For
            End If
        Next keyword
        
        ' Hapus bagian yang tidak diinginkan
        newText = Replace(originalText, "REMNANT kosong; ", "")
        newText = Replace(newText, "THICK2 kosong; ", "")
        newText = Replace(newText, "LENGTH2 kosong; ", "")
        newText = Replace(newText, "WIDTH2 kosong; ", "")
        newText = Replace(newText, "WEIGHT kosong", "")
        
        ' Update sel dengan teks baru
        ws.Cells(i, 16).Value = Trim(newText) ' Trim untuk menghapus spasi di awal/akhir
    Next i
End Sub

