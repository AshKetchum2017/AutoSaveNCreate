Option Explicit

' Class Module (Name) = SNCFolderParser.
' Edit/tambah pasangan di sini; koreksi juga berlaku di dalam kata gabungan.
Private Const TYPO_RULES As String = "prnting=printing;printng=printing;rklame=reklame;rekalme=reklame;reklaem=reklame"

Public Sub ValidateTemplate(ByVal templateText As String)
    Dim i As Long
    Dim ch As String
    Dim inQuotes As Boolean
    Dim hasContent As Boolean
    For i = 1 To Len(templateText)
        ch = Mid$(templateText, i, 1)
        If ch = Chr$(34) Then
            inQuotes = Not inQuotes
        ElseIf inQuotes Then
            If AscW(ch) >= 0 And AscW(ch) < 32 Then Err.Raise 5, "SNCFolderParser", "Pola token tidak boleh mengandung tab atau baris baru."
            If ch <> " " Then hasContent = True
        Else
            Select Case ch
                Case " "
                Case "-", "/", "~"
                    hasContent = True
                Case Else
                    Err.Raise 5, "SNCFolderParser", "Token tidak dikenal pada posisi " & CStr(i) & ": " & ch & vbCrLf & _
                        "Gunakan tanda kutip untuk teks literal. Token __ belum digunakan."
            End Select
        End If
    Next i
    If inQuotes Then Err.Raise 5, "SNCFolderParser", "Tanda kutip pada pola token belum ditutup."
    If Not hasContent Then Err.Raise 5, "SNCFolderParser", "Pola token tidak boleh kosong."
End Sub

Public Function BuildFolderName(ByVal templateText As String, ByVal sourceDirectory As String) As String
    Dim i As Long
    Dim ch As String
    Dim inQuotes As Boolean
    Dim result As String
    Dim category As String
    Dim parentName As String
    Dim fso As Object

    ValidateTemplate templateText
    Set fso = CreateObject("Scripting.FileSystemObject")
    For i = 1 To Len(templateText)
        ch = Mid$(templateText, i, 1)
        If ch = Chr$(34) Then
            inQuotes = Not inQuotes
        ElseIf inQuotes Then
            result = result & ch
        Else
            Select Case ch
                Case " ", "-"
                    result = result & ch
                Case "/"
                    If Len(category) = 0 Then category = CategoryFromPath(sourceDirectory)
                    result = result & category
                Case "~"
                    If Len(parentName) = 0 Then
                        parentName = CorrectParentName(fso.GetFileName(RemoveTrailingSlash(sourceDirectory)))
                        If Len(parentName) = 0 Then Err.Raise 5, "SNCFolderParser", "Nama folder sumber untuk token ~ tidak ditemukan."
                    End If
                    result = result & parentName
                Case Else
                    Err.Raise 5, "SNCFolderParser", "Token tidak dikenal pada posisi " & CStr(i) & ": " & ch & vbCrLf & _
                        "Gunakan tanda kutip untuk teks literal. Token __ belum digunakan."
            End Select
        End If
    Next i
    If inQuotes Then Err.Raise 5, "SNCFolderParser", "Tanda kutip pada pola token belum ditutup."
    ValidateFolderName result
    BuildFolderName = result
End Function

Public Function CategoryFromPath(ByVal sourceDirectory As String) As String
    Dim fso As Object
    Dim currentPath As String
    Dim datePath As String
    Dim monthPath As String
    Dim category As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    currentPath = RemoveTrailingSlash(sourceDirectory)
    Do While Len(currentPath) > 0
        datePath = fso.GetParentFolderName(currentPath)
        monthPath = fso.GetParentFolderName(datePath)
        If IsMonthFolder(fso.GetFileName(monthPath)) And IsDayFolder(fso.GetFileName(datePath)) Then
            category = CategoryCode(fso.GetFileName(currentPath))
            If Len(category) = 0 Then
                Err.Raise 5, "SNCFolderParser", "Kategori folder tidak dikenali: " & fso.GetFileName(currentPath)
            End If
            CategoryFromPath = category
            Exit Function
        End If
        If StrComp(currentPath, datePath, vbTextCompare) = 0 Then Exit Do
        currentPath = datePath
    Loop
    Err.Raise 5, "SNCFolderParser", "Token / tidak menemukan struktur folder bulan, tanggal, dan kategori pada lokasi CDR."
End Function

Private Function CategoryCode(ByVal folderName As String) As String
    Select Case UCase$(folderName)
        Case "RESELLER", "RESELLER CENTER", "RESELER CENTER", "RESELER", "RESELLER CENTRE", "RESELER CENTRE"
            CategoryCode = "rc"
        Case "BLACKPAINT PRINTSHOP", "BLACKPAINT", "BLACK PAINT PRINTSHOP", "BLACK PAINT", "BLACKPAINT PRINT SHOP", "BLACK PAINT PRINT SHOP"
            CategoryCode = "bp"
        Case "FOLDER PRINTEX", "FOLDERPRINTEX", "FOLDERPRINT", "FOLDER PRINT", "FOLDER"
            CategoryCode = "fp"
        Case "CORPORATE"
            CategoryCode = "co"
        Case "CETAK ULANG"
            CategoryCode = "cu"
    End Select
End Function

Private Function IsMonthFolder(ByVal folderName As String) As Boolean
    Dim months As Variant
    Dim i As Long
    months = Split("JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER", "|")
    For i = 0 To UBound(months)
        If StrComp(folderName, Right$("0" & CStr(i + 1), 2) & ". " & months(i), vbTextCompare) = 0 Then
            IsMonthFolder = True
            Exit Function
        End If
    Next i
End Function

' Tanggal berasal dari ancestor CDR, bukan Date/Now atau directory tujuan.
Public Function SourceDateDirectory(ByVal sourceDirectory As String) As String
    Dim fso As Object
    Dim dayPath As String
    Dim monthPath As String
    Dim yearText As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    dayPath = RemoveTrailingSlash(sourceDirectory)
    Do While Len(dayPath) > 0
        monthPath = fso.GetParentFolderName(dayPath)
        If IsDayFolder(fso.GetFileName(dayPath)) And IsMonthFolder(fso.GetFileName(monthPath)) Then
            yearText = SourceYear(fso.GetParentFolderName(monthPath))
            SourceDateDirectory = yearText & "\" & UCase$(fso.GetFileName(monthPath)) & "\" & fso.GetFileName(dayPath)
            Exit Function
        End If
        If StrComp(dayPath, monthPath, vbTextCompare) = 0 Then Exit Do
        dayPath = monthPath
    Loop
    Err.Raise 5, "SNCFolderParser", "Folder bulan/tanggal sumber tidak ditemukan. Gunakan DisposableSave untuk directory khusus."
End Function

Private Function SourceYear(ByVal ancestorPath As String) As String
    Dim fso As Object
    Dim currentPath As String
    Dim folderName As String
    Dim yearText As String
    Dim pass As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' Pass 1: folder tepat empat digit diutamakan di seluruh ancestor.
    ' Pass 2: tahun dalam teks folder, mulai ancestor terdekat bulan.
    For pass = 1 To 2
        currentPath = ancestorPath
        Do While Len(currentPath) > 0
            ' Nama server UNC bukan folder sumber.
            If Left$(currentPath, 2) = "\\" And InStr(3, currentPath, "\") = 0 Then Exit Do
            folderName = fso.GetFileName(currentPath)
            If pass = 1 Then
                If Len(folderName) = 4 And IsFourDigitYear(folderName) Then
                    SourceYear = folderName
                    Exit Function
                End If
            Else
                yearText = YearInFolderName(folderName)
                If Len(yearText) > 0 Then
                    SourceYear = yearText
                    Exit Function
                End If
            End If
            currentPath = fso.GetParentFolderName(currentPath)
        Loop
    Next pass
    Err.Raise 5, "SNCFolderParser", "Tahun sumber tidak ditemukan pada folder tahun maupun nama folder induk. Gunakan DisposableSave."
End Function

Private Function IsFourDigitYear(ByVal value As String) As Boolean
    Dim i As Long
    If Len(value) <> 4 Then Exit Function
    For i = 1 To 4
        If InStr(1, "0123456789", Mid$(value, i, 1), vbBinaryCompare) = 0 Then Exit Function
    Next i
    IsFourDigitYear = (Left$(value, 1) <> "0")
End Function

Private Function YearInFolderName(ByVal folderName As String) As String
    Dim i As Long
    Dim digits As String
    Dim ch As String
    For i = 1 To Len(folderName) + 1
        ch = Mid$(folderName & " ", i, 1)
        If InStr(1, "0123456789", ch, vbBinaryCompare) > 0 Then
            digits = digits & ch
        Else
            If IsFourDigitYear(digits) Then
                If Len(YearInFolderName) > 0 And YearInFolderName <> digits Then
                    Err.Raise 5, "SNCFolderParser", "Tahun ambigu pada folder " & folderName & ". Gunakan DisposableSave."
                End If
                YearInFolderName = digits
            End If
            digits = vbNullString
        End If
    Next i
End Function

Private Function IsDayFolder(ByVal folderName As String) As Boolean
    Dim i As Long
    If Len(folderName) < 1 Or Len(folderName) > 2 Then Exit Function
    For i = 1 To Len(folderName)
        If InStr(1, "0123456789", Mid$(folderName, i, 1), vbBinaryCompare) = 0 Then Exit Function
    Next i
    IsDayFolder = (CLng(folderName) >= 1 And CLng(folderName) <= 31)
End Function

Private Function CorrectParentName(ByVal folderName As String) As String
    Dim rules As Variant
    Dim pair As Variant
    Dim i As Long
    Dim result As String
    result = LCase$(folderName)
    rules = Split(TYPO_RULES, ";")
    For i = 0 To UBound(rules)
        pair = Split(rules(i), "=")
        result = Replace(result, pair(0), pair(1), 1, -1, vbBinaryCompare)
    Next i
    CorrectParentName = result
End Function

Private Function RemoveTrailingSlash(ByVal folderPath As String) As String
    Do While Len(folderPath) > 3 And Right$(folderPath, 1) = "\"
        folderPath = Left$(folderPath, Len(folderPath) - 1)
    Loop
    RemoveTrailingSlash = folderPath
End Function

Public Sub ValidateFolderName(ByVal folderName As String)
    Dim i As Long
    Dim ch As String
    Dim stem As String
    Dim dotPosition As Long

    If Len(Trim$(folderName)) = 0 Then Err.Raise 5, "SNCFolderParser", "Nama folder hasil parsing kosong."
    If Len(folderName) > 255 Then Err.Raise 5, "SNCFolderParser", "Nama folder hasil parsing melebihi 255 karakter."
    If Right$(folderName, 1) = " " Or Right$(folderName, 1) = "." Then
        Err.Raise 5, "SNCFolderParser", "Nama folder tidak boleh diakhiri spasi atau titik."
    End If
    For i = 1 To Len(folderName)
        ch = Mid$(folderName, i, 1)
        If InStr(1, "\/:*?" & Chr$(34) & "<>|", ch, vbBinaryCompare) > 0 Or AscW(ch) < 32 And AscW(ch) >= 0 Then
            Err.Raise 5, "SNCFolderParser", "Nama folder mengandung karakter yang tidak diizinkan pada posisi " & CStr(i) & "."
        End If
    Next i
    stem = UCase$(folderName)
    dotPosition = InStr(1, stem, ".", vbBinaryCompare)
    If dotPosition > 0 Then stem = Left$(stem, dotPosition - 1)
    stem = RTrim$(stem)
    Select Case stem
        Case "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", _
             "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
            Err.Raise 5, "SNCFolderParser", "Nama folder merupakan nama perangkat Windows yang dicadangkan: " & folderName
    End Select
End Sub
