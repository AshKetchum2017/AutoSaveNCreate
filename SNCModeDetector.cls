Option Explicit

' Class Module (Name) = SNCModeDetector. Edit daftar keyword per kategori di sini.
Private Const KISSA_KEYWORDS As String = "KISS CUT|CUTTING BAHAN|LASER|VINYL|CHROMO|A3|A3+"
Private Const DIEA_KEYWORDS As String = "DIE CUT|CUTTING BAHAN|LASER|KARTU NAMA|GR|KOTAK|VINYL|CHROMO|A3|A3+|LONG|EXT|EXTENDED|AP|ART PAPER|APA 3|APA 3+"
Private Const HIKISS_KEYWORDS As String = "KISS CUT|ECO SOLVENT|CUTTING BAHAN|EPSON|UV PREMIUM|UV GEL|METERAN|RITRAMA|MAXDECAL|ESSENS|HIGHRESS"
Private Const HIDIE_KEYWORDS As String = "DIE CUT|POTONG CUTTER|ECO SOLVENT|EPSON|UV PREMIUM|UV GEL|METERAN|RITRAMA|MAXDECAL|ESSENS|INFRABOARD|FOAMBOARD|HIGHRESS"
Private Const KISS_MARKERS As String = "KISS CUT|KIS CUT|KISSCUT|KISS"
Private Const DIE_MARKERS As String = "DIE CUT|DIECUT|DIE CUTT|DIECUTT|POT CUTTER|POT. CUTTER|POTONG CUTTER"
Private Const CUTTER_MARKERS As String = "POTONG CUTTER|POT CUTTER|CUTTER"
Private Const HI_SHORT_KEYWORDS As String = "ECO|EPS|UV"
Private Const DIRECTORY_DIE_MARKERS As String = "DIE"

' Directory lebih dahulu; nama CDR hanya melengkapi petunjuk yang belum lengkap.
' Hasil kosong + conflict=True membutuhkan keputusan Lanjut/Cancel dari form.
Public Function DetectDirectoryMode(ByVal folderPath As String, ByVal fileName As String, ByRef conflict As Boolean) As String
    Dim fso As Object
    Dim currentPath As String
    Dim parentPath As String
    Dim words As Variant
    Dim kiss As Boolean, die As Boolean, smallFormat As Boolean, highFormat As Boolean
    Dim fileKiss As Boolean, fileDie As Boolean, fileSmall As Boolean, fileHigh As Boolean
    conflict = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    currentPath = folderPath
    Do While Len(currentPath) > 0
        ReadSignals fso.GetFileName(currentPath), kiss, die, smallFormat, highFormat
        words = Split(NormalizeWords(fso.GetFileName(currentPath)), " ")
        die = die Or HasAny(words, DIRECTORY_DIE_MARKERS)
        parentPath = fso.GetParentFolderName(currentPath)
        If StrComp(parentPath, currentPath, vbTextCompare) = 0 Then Exit Do
        currentPath = parentPath
    Loop
    If Not (kiss Or die Or smallFormat Or highFormat) Then Exit Function
    If (kiss And die) Or (smallFormat And highFormat) Then
        conflict = True
        Exit Function
    End If
    If Not ((kiss Or die) And (smallFormat Or highFormat)) Then
        ReadSignals fso.GetBaseName(fileName), fileKiss, fileDie, fileSmall, fileHigh
        kiss = kiss Or fileKiss
        die = die Or fileDie
        smallFormat = smallFormat Or fileSmall
        highFormat = highFormat Or fileHigh
    End If
    If (kiss And die) Or (smallFormat And highFormat) Then
        conflict = True
        Exit Function
    End If
    DetectDirectoryMode = ModeFromSignals(kiss, die, smallFormat, highFormat)
End Function

Public Function ModeCaption(ByVal modeName As String) As String
    Select Case modeName
        Case "DieA": ModeCaption = "Die Cut A3"
        Case "KissA": ModeCaption = "Kiss Cut A3"
        Case "HiDie": ModeCaption = "Hires Die Cut"
        Case "HiKiss": ModeCaption = "Hires Kiss Cut"
    End Select
End Function

Private Sub ReadSignals(ByVal value As String, ByRef kiss As Boolean, ByRef die As Boolean, ByRef smallFormat As Boolean, ByRef highFormat As Boolean)
    Dim words As Variant
    Dim cutter As Boolean
    words = Split(NormalizeWords(value), " ")
    kiss = kiss Or HasAny(words, KISS_MARKERS)
    cutter = HasAny(words, CUTTER_MARKERS)
    die = die Or HasAny(words, DIE_MARKERS) Or cutter
    smallFormat = smallFormat Or HasGroup(words, KISSA_KEYWORDS) Or HasGroup(words, DIEA_KEYWORDS)
    highFormat = highFormat Or HasGroup(words, HIKISS_KEYWORDS) Or HasGroup(words, HIDIE_KEYWORDS) Or HasAny(words, HI_SHORT_KEYWORDS) Or cutter
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
    Set fso = CreateObject("Scripting.FileSystemObject")
    ReadSignals fso.GetBaseName(fileName), kiss, die, smallFormat, highFormat
    DetectMode = ModeFromSignals(kiss, die, smallFormat, highFormat)
End Function

Private Function HasGroup(ByRef words As Variant, ByVal keywords As String) As Boolean
    Dim items As Variant
    Dim i As Long
    items = Split(keywords, "|")
    For i = 0 To UBound(items)
        Select Case CStr(items(i))
            Case "KISS CUT", "DIE CUT", "POTONG CUTTER"
                ' Penanda jenis potong tidak sekaligus membuktikan kelompok produksi.
            Case Else
                If MatchesPhrase(words, CStr(items(i))) Then
                    HasGroup = True
                    Exit Function
                End If
        End Select
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
