Option Explicit

' Class Module (Name) = SNCModeDetector. Keyword mode langsung didahulukan dari petunjuk umum.
Private Const A_MACHINE_KEYWORDS As String = "LASER"
Private Const HI_MACHINE_KEYWORDS As String = "ECO|ECO SOLVENT|UV PREMIUM|UV GEL|EPSON"
Private Const A_BRAND_KEYWORDS As String = "QUANTAC|K-TAK"
Private Const HI_BRAND_KEYWORDS As String = "ESSENS|HCR|RITRAMA|MAXDECAL"
Private Const SHARED_BRAND_KEYWORDS As String = "GRAFTAC"
Private Const A_MATERIAL_KEYWORDS As String = "VINYL|CROMO|CHROMO|ART PAPER|AP|GOLD|SILVER|LMO|KRAFT"
Private Const HI_MATERIAL_KEYWORDS As String = "CHROME|SANDBLAST"
' Keyword ini berlaku untuk keempat finishing tanpa menentukan mode sendiri.
Private Const ALL_FINISHING_KEYWORDS As String = "CUTTING BAHAN"
Private Const DIRECT_DIEA_KEYWORDS As String = "KARTU NAMA|LONG|EXTENDED|EXT"
Private Const DIRECT_HIDIE_KEYWORDS As String = "INFRABOARD|FOAMBOARD"
' Jenis bahan dan pilihan berikut tidak menentukan A/Hi sendiri.
Private Const SHARED_MATERIAL_TYPE_KEYWORDS As String = "TRANSPARANT|TRANSPARENT|GLOSSY|DOFF|MATTE|MATE|HOLOGRAM"
Private Const UV_OPTION_KEYWORDS As String = "WHITE"
Private Const MAXDECAL_OPTION_KEYWORDS As String = "PREMIUM|LMP|LABEL"
Private Const KISS_MARKERS As String = "KISS CUT|KIS CUT|KISSCUT|KISS|KIS|KISS CUTT|KISSCUTT|MASKING"
Private Const DIE_MARKERS As String = "DIE CUT|DIECUT|DIE CUTT|DIECUTT|DIE"
Private Const CUTTER_MARKERS As String = "POTONG CUTTER|POT CUTTER|POTONG|CUTTER|POT. CUTTER|POT|POT."
Private Const DIRECTORY_DIE_MARKERS As String = "DIE"

' Directory lebih dahulu; nama CDR hanya melengkapi petunjuk yang belum lengkap.
' Hasil kosong + conflict=True membutuhkan keputusan Lanjut/Cancel dari form.
Public Function DetectDirectoryMode(ByVal folderPath As String, ByVal fileName As String, ByRef conflict As Boolean) As String
    Dim fso As Object
    Dim currentPath As String
    Dim parentPath As String
    Dim words As Variant
    Dim kiss As Boolean, die As Boolean, smallFormat As Boolean, highFormat As Boolean, graftac As Boolean
    Dim fileKiss As Boolean, fileDie As Boolean, fileSmall As Boolean, fileHigh As Boolean, fileGraftac As Boolean
    Dim directDieA As Boolean, directHiDie As Boolean, fileDirectDieA As Boolean, fileDirectHiDie As Boolean
    Dim effectiveHigh As Boolean
    conflict = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    currentPath = folderPath
    Do While Len(currentPath) > 0
        ReadSignals fso.GetFileName(currentPath), kiss, die, smallFormat, highFormat, graftac, directDieA, directHiDie
        words = Split(NormalizeWords(fso.GetFileName(currentPath)), " ")
        die = die Or HasAny(words, DIRECTORY_DIE_MARKERS)
        parentPath = fso.GetParentFolderName(currentPath)
        If StrComp(parentPath, currentPath, vbTextCompare) = 0 Then Exit Do
        currentPath = parentPath
    Loop
    If directDieA And directHiDie Then
        conflict = True
        Exit Function
    End If
    If directDieA Then
        DetectDirectoryMode = "DieA"
        Exit Function
    End If
    If directHiDie Then
        DetectDirectoryMode = "HiDie"
        Exit Function
    End If
    effectiveHigh = highFormat Or (graftac And Not smallFormat)
    If Not (kiss Or die Or smallFormat Or effectiveHigh) Then Exit Function
    If (kiss And die) Or (smallFormat And effectiveHigh) Then
        conflict = True
        Exit Function
    End If
    If Not ((kiss Or die) And (smallFormat Or effectiveHigh)) Then
        ReadSignals fso.GetBaseName(fileName), fileKiss, fileDie, fileSmall, fileHigh, fileGraftac, fileDirectDieA, fileDirectHiDie
        If fileDirectDieA And fileDirectHiDie Then
            conflict = True
            Exit Function
        End If
        If fileDirectDieA Then
            DetectDirectoryMode = "DieA"
            Exit Function
        End If
        If fileDirectHiDie Then
            DetectDirectoryMode = "HiDie"
            Exit Function
        End If
        kiss = kiss Or fileKiss
        die = die Or fileDie
        smallFormat = smallFormat Or fileSmall
        highFormat = highFormat Or fileHigh
        graftac = graftac Or fileGraftac
    End If
    effectiveHigh = highFormat Or (graftac And Not smallFormat)
    If (kiss And die) Or (smallFormat And effectiveHigh) Then
        conflict = True
        Exit Function
    End If
    DetectDirectoryMode = ModeFromSignals(kiss, die, smallFormat, effectiveHigh)
End Function

Public Function ModeCaption(ByVal modeName As String) As String
    Select Case modeName
        Case "DieA": ModeCaption = "Die Cut A3"
        Case "KissA": ModeCaption = "Kiss Cut A3"
        Case "HiDie": ModeCaption = "Hires Die Cut"
        Case "HiKiss": ModeCaption = "Hires Kiss Cut"
    End Select
End Function

Private Sub ReadSignals(ByVal value As String, ByRef kiss As Boolean, ByRef die As Boolean, ByRef smallFormat As Boolean, ByRef highFormat As Boolean, ByRef graftac As Boolean, ByRef directDieA As Boolean, ByRef directHiDie As Boolean)
    Dim words As Variant
    Dim cutter As Boolean
    Dim aKeywords As String
    Dim hiKeywords As String
    words = Split(NormalizeWords(value), " ")
    kiss = kiss Or HasAny(words, KISS_MARKERS)
    cutter = HasAny(words, CUTTER_MARKERS)
    die = die Or HasAny(words, DIE_MARKERS) Or cutter
    aKeywords = A_MACHINE_KEYWORDS & "|" & A_BRAND_KEYWORDS & "|" & A_MATERIAL_KEYWORDS
    hiKeywords = HI_MACHINE_KEYWORDS & "|" & HI_BRAND_KEYWORDS & "|" & HI_MATERIAL_KEYWORDS
    smallFormat = smallFormat Or HasCategory(words, aKeywords)
    highFormat = highFormat Or HasCategory(words, hiKeywords)
    graftac = graftac Or HasAny(words, SHARED_BRAND_KEYWORDS)
    ' Mode langsung memakai frasa exact agar typo pada kata umum tidak menimpa petunjuk lain.
    directDieA = directDieA Or HasAnyExact(words, DIRECT_DIEA_KEYWORDS)
    directHiDie = directHiDie Or HasAnyExact(words, DIRECT_HIDIE_KEYWORDS)
End Sub

Private Function ModeFromSignals(ByVal kiss As Boolean, ByVal die As Boolean, ByVal smallFormat As Boolean, ByVal highFormat As Boolean) As String
    If kiss = die Or smallFormat = highFormat Then Exit Function
    If smallFormat Then
        If kiss Then ModeFromSignals = "KissA" Else ModeFromSignals = "DieA"
    Else
        If kiss Then ModeFromSignals = "HiKiss" Else ModeFromSignals = "HiDie"
    End If
End Function

' Hasil kosong: kurang petunjuk/bertentangan, pemanggil memulihkan default UF.
Public Function DetectMode(ByVal fileName As String) As String
    Dim fso As Object
    Dim kiss As Boolean
    Dim die As Boolean
    Dim smallFormat As Boolean
    Dim highFormat As Boolean
    Dim graftac As Boolean
    Dim directDieA As Boolean
    Dim directHiDie As Boolean
    Set fso = CreateObject("Scripting.FileSystemObject")
    ReadSignals fso.GetBaseName(fileName), kiss, die, smallFormat, highFormat, graftac, directDieA, directHiDie
    If directDieA And directHiDie Then Exit Function
    If directDieA Then
        DetectMode = "DieA"
        Exit Function
    End If
    If directHiDie Then
        DetectMode = "HiDie"
        Exit Function
    End If
    DetectMode = ModeFromSignals(kiss, die, smallFormat, highFormat Or (graftac And Not smallFormat))
End Function

Private Function HasAnyExact(ByRef words As Variant, ByVal keywords As String) As Boolean
    Dim items As Variant
    Dim i As Long
    Dim exactWords As String
    exactWords = " " & Join(words, " ") & " "
    items = Split(keywords, "|")
    For i = 0 To UBound(items)
        If InStr(1, exactWords, " " & NormalizeWords(CStr(items(i))) & " ", vbBinaryCompare) > 0 Then
            HasAnyExact = True
            Exit Function
        End If
    Next i
End Function

Private Function HasCategory(ByRef words As Variant, ByVal keywords As String) As Boolean
    Dim items As Variant
    Dim i As Long
    Dim keyword As String
    Dim exactWords As String
    items = Split(keywords, "|")
    exactWords = " " & Join(words, " ") & " "
    For i = 0 To UBound(items)
        keyword = NormalizeWords(CStr(items(i)))
        ' CHROMO dan CHROME beda kelompok; jangan samakan lewat toleransi typo.
        If keyword = "CHROMO" Or keyword = "CHROME" Then
            If InStr(1, exactWords, " " & keyword & " ", vbBinaryCompare) > 0 Then
                HasCategory = True
                Exit Function
            End If
        ElseIf MatchesPhrase(words, keyword) Then
            HasCategory = True
            Exit Function
        End If
    Next i
End Function

Private Function HasAny(ByRef words As Variant, ByVal keywords As String) As Boolean
    Dim items As Variant
    Dim i As Long
    items = Split(keywords, "|")
    For i = 0 To UBound(items)
        If MatchesPhrase(words, CStr(items(i))) Then
            HasAny = True
            Exit Function
        End If
    Next i
End Function

Private Function MatchesPhrase(ByRef words As Variant, ByVal keyword As String) As Boolean
    Dim expected As Variant
    Dim i As Long
    Dim j As Long
    Dim matches As Boolean
    expected = Split(NormalizeWords(keyword), " ")
    For i = 0 To UBound(words) - UBound(expected)
        matches = True
        For j = 0 To UBound(expected)
            If Not WordMatches(CStr(words(i + j)), CStr(expected(j))) Then
                matches = False
                Exit For
            End If
        Next j
        If matches Then
            MatchesPhrase = True
            Exit Function
        End If
    Next i
End Function

Private Function NormalizeWords(ByVal value As String) As String
    Dim i As Long
    Dim ch As String
    Dim result As String
    value = UCase$(value)
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789+", ch, vbBinaryCompare) > 0 Then
            result = result & ch
        ElseIf Len(result) > 0 Then
            If Right$(result, 1) <> " " Then result = result & " "
        End If
    Next i
    NormalizeWords = Trim$(result)
End Function

' Kata pendek harus exact/alias eksplisit. Kata >=4 huruf boleh satu edit
' (tambah/hapus/ganti huruf) atau pertukaran dua huruf bersebelahan.
Private Function WordMatches(ByVal actual As String, ByVal expected As String) As Boolean
    Dim i As Long
    Dim j As Long
    Dim mismatch As Long
    Dim firstMismatch As Long
    Dim swapText As String
    If actual = expected Then
        WordMatches = True
        Exit Function
    End If
    If Len(actual) < 4 Or Len(expected) < 4 Then Exit Function
    If Abs(Len(actual) - Len(expected)) > 1 Then Exit Function
    If Len(actual) = Len(expected) Then
        For i = 1 To Len(actual)
            If Mid$(actual, i, 1) <> Mid$(expected, i, 1) Then
                mismatch = mismatch + 1
                If firstMismatch = 0 Then firstMismatch = i
            End If
        Next i
        If mismatch = 1 Then
            WordMatches = True
        ElseIf mismatch = 2 And firstMismatch < Len(actual) Then
            swapText = Left$(actual, firstMismatch - 1) & Mid$(actual, firstMismatch + 1, 1) & _
                Mid$(actual, firstMismatch, 1) & Mid$(actual, firstMismatch + 2)
            WordMatches = (swapText = expected)
        End If
        Exit Function
    End If
    If Len(actual) < Len(expected) Then
        swapText = actual
        actual = expected
        expected = swapText
    End If
    i = 1
    j = 1
    Do While i <= Len(actual) And j <= Len(expected)
        If Mid$(actual, i, 1) = Mid$(expected, j, 1) Then
            j = j + 1
        Else
            mismatch = mismatch + 1
            If mismatch > 1 Then Exit Function
        End If
        i = i + 1
    Loop
    WordMatches = True
End Function
