Option Explicit

' Class Module (Name) = SNCSaveRunner.
' Hasil kosong berarti user membatalkan. Error diteruskan beserta operasi/path.
Public Function SaveDocument(ByVal doc As Document, ByVal baseDirectory As String, ByVal tokenText As String, _
                             Optional ByVal disposableDirectory As Boolean = False, _
                             Optional ByVal corelVersionLabel As String = "", _
                             Optional ByVal embedColorProfiles As Boolean = True) As String
    Dim fso As Object
    Dim parser As SNCFolderParser
    Dim options As StructSaveAsOptions
    Dim fileName As String
    Dim folderName As String
    Dim targetFolder As String
    Dim targetPath As String
    Dim overwrite As Boolean
    Dim answer As VbMsgBoxResult
    Dim operation As String
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim dateDirectory As String
    Dim settings As SNCSettingsStore
    Dim targetVersion As cdrFileVersion

    On Error GoTo SaveFailed
    operation = "Memeriksa dokumen sumber"
    If doc Is Nothing Then Err.Raise 5, "SNCSaveRunner", "Tidak ada dokumen aktif."
    If Len(doc.FileName) = 0 Or Len(doc.FilePath) = 0 Then
        Err.Raise 5, "SNCSaveRunner", "Simpan dokumen sumber sebagai CDR terlebih dahulu agar nama dan folder sumber tersedia."
    End If
    operation = "Memvalidasi versi output CDR"
    Set settings = New SNCSettingsStore
    If Len(corelVersionLabel) = 0 Then corelVersionLabel = settings.LoadCorelVersion()
    targetVersion = settings.ResolveCorelFileVersion(corelVersionLabel)
    Set fso = CreateObject("Scripting.FileSystemObject")
    fileName = fso.GetFileName(doc.FileName)
    If StrComp(fso.GetExtensionName(fileName), "cdr", vbTextCompare) <> 0 Then
        Err.Raise 5, "SNCSaveRunner", "Dokumen sumber harus berupa file CDR."
    End If
    operation = "Memeriksa directory tujuan"
    If Not fso.FolderExists(baseDirectory) Then Err.Raise 5, "SNCSaveRunner", "Atur directory tujuan mode yang dipilih terlebih dahulu."
    Set parser = New SNCFolderParser
    operation = "Parsing nama folder"
    folderName = parser.BuildFolderName(tokenText, doc.FilePath)
    If disposableDirectory Then
        targetFolder = fso.BuildPath(baseDirectory, folderName)
    Else
        operation = "Membaca tahun, bulan, dan tanggal sumber"
        dateDirectory = parser.SourceDateDirectory(doc.FilePath)
        targetFolder = fso.BuildPath(fso.BuildPath(baseDirectory, dateDirectory), folderName)
    End If
    If fso.FileExists(targetFolder) Then Err.Raise 5, "SNCSaveRunner", "Nama folder tujuan sudah digunakan oleh sebuah file: " & targetFolder
    targetPath = fso.BuildPath(targetFolder, fileName)
    operation = "Memeriksa file tujuan"
    If fso.FolderExists(targetPath) Then Err.Raise 5, "SNCSaveRunner", "Nama file tujuan sudah digunakan oleh sebuah folder: " & targetPath
    If fso.FileExists(targetPath) Then
        answer = MsgBox("File CDR tujuan sudah ada:" & vbCrLf & targetPath & vbCrLf & vbCrLf & _
            "Yes = buat duplikat dengan nomor di belakang nama." & vbCrLf & _
            "No = timpa file CDR existing." & vbCrLf & _
            "Cancel = batalkan penyimpanan.", vbYesNoCancel Or vbQuestion Or vbDefaultButton3, "AutoSaveNCreate")
        Select Case answer
            Case vbYes
                targetPath = NextDuplicatePath(targetFolder, fileName)
            Case vbNo
                overwrite = True
            Case Else
                Exit Function
        End Select
    End If
    operation = "Menyiapkan opsi Save As CDR"
    Set options = Application.CreateStructSaveAsOptions
    options.Filter = cdrCDR
    options.Range = cdrAllPages
    options.Version = targetVersion
    options.EmbedICCProfile = embedColorProfiles
    options.EmbedVBAProject = True
    options.Overwrite = overwrite
    operation = "Membuat folder tujuan"
    EnsureTargetFolder baseDirectory, targetFolder
    operation = "Document.SaveAs"
    doc.SaveAs targetPath, options
    operation = "Memeriksa hasil Save As"
    If Not fso.FileExists(targetPath) Then Err.Raise 5, "SNCSaveRunner", "File hasil Save As tidak ditemukan."
    If fso.GetFile(targetPath).Size = 0 Then Err.Raise 5, "SNCSaveRunner", "File hasil Save As kosong."
    If StrComp(fso.GetAbsolutePathName(doc.FullFileName), fso.GetAbsolutePathName(targetPath), vbTextCompare) <> 0 Then
        Err.Raise 5, "SNCSaveRunner", "Dokumen belum mengikuti lokasi tujuan setelah Save As."
    End If
    SaveDocument = targetPath
    Exit Function
SaveFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Err.Raise errorNumber, "SNCSaveRunner", "Operasi: " & operation & vbCrLf & _
        "Tujuan: " & targetPath & vbCrLf & errorDescription
End Function

Private Sub EnsureTargetFolder(ByVal baseDirectory As String, ByVal targetFolder As String)
    Dim fso As Object
    Dim pending As Collection
    Dim currentPath As String
    Dim parentPath As String
    Dim rootPath As String
    Dim i As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set pending = New Collection
    rootPath = fso.GetAbsolutePathName(baseDirectory)
    If Right$(rootPath, 1) <> "\" Then rootPath = rootPath & "\"
    currentPath = fso.GetAbsolutePathName(targetFolder)
    If StrComp(Left$(currentPath, Len(rootPath)), rootPath, vbTextCompare) <> 0 Then
        Err.Raise 5, "SNCSaveRunner", "Folder tujuan berada di luar directory yang dipilih."
    End If
    Do While Not fso.FolderExists(currentPath)
        If fso.FileExists(currentPath) Then Err.Raise 5, "SNCSaveRunner", "Nama folder sudah digunakan file: " & currentPath
        pending.Add currentPath
        parentPath = fso.GetParentFolderName(currentPath)
        If Len(parentPath) < Len(rootPath) - 1 Then Err.Raise 5, "SNCSaveRunner", "Root directory tujuan tidak lagi tersedia."
        currentPath = parentPath
    Loop
    For i = pending.Count To 1 Step -1
        fso.CreateFolder CStr(pending(i))
    Next i
End Sub

Public Function NextDuplicatePath(ByVal targetFolder As String, ByVal fileName As String) As String
    Dim fso As Object
    Dim candidate As String
    Dim suffix As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    suffix = 1
    Do
        candidate = fso.BuildPath(targetFolder, fso.GetBaseName(fileName) & " (" & CStr(suffix) & ")." & fso.GetExtensionName(fileName))
        If Not fso.FileExists(candidate) And Not fso.FolderExists(candidate) Then
            NextDuplicatePath = candidate
            Exit Function
        End If
        suffix = suffix + 1
    Loop
End Function
