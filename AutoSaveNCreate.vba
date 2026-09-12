Option Explicit

Private pUsers As Collection
Private pProcessing As Boolean
Private pDirectorySession As SNCDirectorySession
Private pModeDocument As Document
Private pDefaultModes As Variant
Private pLoadingCorelVersion As Boolean
Private pLoadingEmbedding As Boolean

Private Sub UserForm_Initialize()
    On Error GoTo LoadFailed
    ' Ambil nilai designer sebelum deteksi pertama mengubah OptionButton.
    pDefaultModes = Array(CBool(optDieA.Value), CBool(optHiDie.Value), CBool(optKissA.Value), CBool(optHiKiss.Value))
    Set pDirectorySession = GetSNCDirectorySession()
    LoadCorelVersions
    LoadEmbeddingSettings
    LoadUserChoices
    SyncActiveDocument
    Exit Sub
LoadFailed:
    cmdProcess.Enabled = False
    MsgBox "Gagal memuat pengaturan AutoSNC (" & CStr(Err.Number) & "): " & Err.Description, vbExclamation, "AutoSaveNCreate"
End Sub

Private Sub UserForm_Activate()
    On Error GoTo DetectFailed
    If pProcessing Then Exit Sub
    SyncActiveDocument
    Exit Sub
DetectFailed:
    RestoreDefaultMode
    MsgBox "Gagal mendeteksi mode (" & CStr(Err.Number) & "): " & Err.Description, vbExclamation, "AutoSaveNCreate"
End Sub

Private Sub cmbCorelVersion_Change()
    Dim settings As SNCSettingsStore
    If pLoadingCorelVersion Then Exit Sub
    If cmbCorelVersion.ListIndex < 0 Then Exit Sub
    On Error GoTo SaveFailed
    Set settings = New SNCSettingsStore
    settings.SaveCorelVersion CStr(cmbCorelVersion.Value)
    Exit Sub
SaveFailed:
    MsgBox "Gagal menyimpan pilihan versi CDR (" & CStr(Err.Number) & "): " & Err.Description, vbExclamation, "AutoSaveNCreate"
End Sub

Private Sub chkEmbedColorProfiles_Click()
    Dim settings As SNCSettingsStore
    If pLoadingEmbedding Then Exit Sub
    On Error GoTo SaveFailed
    Set settings = New SNCSettingsStore
    settings.SaveEmbedColorProfiles CBool(chkEmbedColorProfiles.Value)
    Exit Sub
SaveFailed:
    MsgBox "Gagal menyimpan Embed Color Profiles (" & CStr(Err.Number) & "): " & Err.Description, vbExclamation, "AutoSaveNCreate"
End Sub

Private Sub chkEmbedFonts_Click()
    Dim settings As SNCSettingsStore
    If pLoadingEmbedding Then Exit Sub
    On Error GoTo SaveFailed
    Set settings = New SNCSettingsStore
    settings.SaveEmbedFonts CBool(chkEmbedFonts.Value)
    MsgBox "Pilihan Embed Fonts sudah disimpan sebagai preferensi." & vbCrLf & _
        "Penerapannya ke Save As CDR belum tersedia; masih menunggu pemetaan API.", vbInformation, "AutoSaveNCreate"
    Exit Sub
SaveFailed:
    MsgBox "Gagal menyimpan Embed Fonts (" & CStr(Err.Number) & "): " & Err.Description, vbExclamation, "AutoSaveNCreate"
End Sub

Private Sub LoadEmbeddingSettings()
    Dim settings As SNCSettingsStore
    Dim errorNumber As Long
    Dim errorDescription As String
    On Error GoTo LoadFailed
    pLoadingEmbedding = True
    Set settings = New SNCSettingsStore
    chkEmbedColorProfiles.Value = settings.LoadEmbedColorProfiles(CBool(chkEmbedColorProfiles.Value))
    chkEmbedFonts.Value = settings.LoadEmbedFonts(CBool(chkEmbedFonts.Value))
    chkEmbedFonts.ControlTipText = "Preferensi tersimpan; penerapan Embed Fonts ke CDR belum tersedia."
    pLoadingEmbedding = False
    Exit Sub
LoadFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    pLoadingEmbedding = False
    Err.Raise errorNumber, "AutoSaveNCreate.LoadEmbeddingSettings", errorDescription
End Sub

Private Sub LoadCorelVersions()
    Dim settings As SNCSettingsStore
    Dim labels As Variant
    Dim savedVersion As String
    Dim i As Long
    Dim errorNumber As Long
    Dim errorDescription As String
    On Error GoTo LoadFailed
    pLoadingCorelVersion = True
    Set settings = New SNCSettingsStore
    labels = settings.CorelVersionLabels()
    savedVersion = settings.LoadCorelVersion()
    cmbCorelVersion.Clear
    cmbCorelVersion.Style = fmStyleDropDownList
    For i = 0 To UBound(labels)
        cmbCorelVersion.AddItem CStr(labels(i))
        If CStr(labels(i)) = savedVersion Then cmbCorelVersion.ListIndex = i
    Next i
    pLoadingCorelVersion = False
    Exit Sub
LoadFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    pLoadingCorelVersion = False
    Err.Raise errorNumber, "AutoSaveNCreate.LoadCorelVersions", errorDescription
End Sub

Private Sub cmbUserSelection_Change()

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDieASetting_Click()
    ShowDirectorySettings "DieA"
End Sub

Private Sub cmdHiDieSetting_Click()
    ShowDirectorySettings "HiDie"
End Sub

Private Sub cmdHiKissSetting_Click()
    ShowDirectorySettings "HiKiss"
End Sub

Private Sub cmdKissASetting_Click()
    ShowDirectorySettings "KissA"
End Sub

Private Sub cmdProcess_Click()
    Dim settings As SNCSettingsStore
    Dim runner As SNCSaveRunner
    Dim doc As Document
    Dim item As Variant
    Dim modeName As String
    Dim savedPath As String
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim baseDirectory As String
    Dim disposable As Boolean
    If pProcessing Then Exit Sub
    On Error GoTo ProcessFailed
    If Application.Documents.Count = 0 Then Err.Raise 5, "AutoSaveNCreate", "Tidak ada dokumen aktif."
    If SyncActiveDocument() Then
        MsgBox "Dokumen aktif berubah. Periksa pilihan mode dan setting user, lalu tekan Process kembali.", vbInformation, "AutoSaveNCreate"
        Exit Sub
    End If
    If pUsers Is Nothing Then Err.Raise 5, "AutoSaveNCreate", "Daftar user belum tersedia."
    If cmbUserSelection.ListIndex < 0 Then Err.Raise 5, "AutoSaveNCreate", "Pilih setting user terlebih dahulu."
    If cmbCorelVersion.ListIndex < 0 Then Err.Raise 5, "AutoSaveNCreate", "Pilih versi output CDR terlebih dahulu."
    modeName = SelectedMode()
    item = pUsers(cmbUserSelection.ListIndex + 1)
    Set doc = ActiveDocument
    Set settings = New SNCSettingsStore
    Set runner = New SNCSaveRunner
    baseDirectory = pDirectorySession.GetDirectory(doc, modeName)
    disposable = (Len(baseDirectory) > 0)
    If Not disposable Then baseDirectory = settings.LoadDirectory(modeName)
    pProcessing = True
    cmdProcess.Enabled = False
    savedPath = runner.SaveDocument(doc, baseDirectory, CStr(item(1)), disposable, _
        CStr(cmbCorelVersion.Value), CBool(chkEmbedColorProfiles.Value))
    If Len(savedPath) > 0 And disposable Then pDirectorySession.ClearDirectory doc, modeName
    pProcessing = False
    cmdProcess.Enabled = True
    If Len(savedPath) > 0 Then MsgBox "CDR berhasil disimpan:" & vbCrLf & savedPath, vbInformation, "AutoSaveNCreate"
    Exit Sub
ProcessFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    pProcessing = False
    cmdProcess.Enabled = True
    MsgBox "Gagal memproses CDR (" & CStr(errorNumber) & "): " & vbCrLf & errorDescription, vbExclamation, "AutoSaveNCreate"
End Sub

Private Sub optDieA_Click()

End Sub

Private Sub optHiDie_Click()

End Sub

Private Sub optHiKiss_Click()

End Sub

Private Sub optKissA_Click()

End Sub

Private Sub cmdUserSettings_Click()
    Dim editor As Object
    Dim errorNumber As Long
    Dim errorDescription As String
    On Error GoTo OpenFailed
    Set editor = UserForms.Add("UserSettingsMenu")
    editor.Show vbModal
    LoadUserChoices
    Exit Sub
OpenFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    If Not editor Is Nothing Then Unload editor
    On Error GoTo 0
    MsgBox "Gagal membuka/memuat setting user (" & CStr(errorNumber) & "): " & errorDescription, vbExclamation, "AutoSaveNCreate"
End Sub

Private Sub LoadUserChoices()
    Dim settings As SNCSettingsStore
    Dim i As Long
    Dim item As Variant
    Dim previousName As String
    If cmbUserSelection.ListIndex >= 0 Then previousName = CStr(cmbUserSelection.Value)
    Set settings = New SNCSettingsStore
    Set pUsers = settings.LoadUsers()
    cmbUserSelection.Clear
    For i = 1 To pUsers.Count
        item = pUsers(i)
        cmbUserSelection.AddItem CStr(item(0))
        If StrComp(CStr(item(0)), previousName, vbTextCompare) = 0 Then cmbUserSelection.ListIndex = i - 1
    Next i
    If cmbUserSelection.ListIndex < 0 And pUsers.Count > 0 Then cmbUserSelection.ListIndex = 0
    cmdProcess.Enabled = (pUsers.Count > 0)
End Sub

Private Function SelectedMode() As String
    Dim count As Long
    If optDieA.Value Then
        SelectedMode = "DieA"
        count = count + 1
    End If
    If optHiDie.Value Then
        SelectedMode = "HiDie"
        count = count + 1
    End If
    If optKissA.Value Then
        SelectedMode = "KissA"
        count = count + 1
    End If
    If optHiKiss.Value Then
        SelectedMode = "HiKiss"
        count = count + 1
    End If
    If count <> 1 Then Err.Raise 5, "AutoSaveNCreate", "Pilih tepat satu mode: DieA, HiDie, KissA, atau HiKiss."
End Function

Private Sub ShowDirectorySettings(ByVal modeName As String)
    Dim editor As Object
    Dim operation As String
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim sourceDocument As Document
    On Error GoTo OpenFailed
    operation = "Membuka DirectorySettings"
    Set editor = UserForms.Add("DirectorySettings")
    operation = "Memuat directory " & modeName
    If Application.Documents.Count > 0 Then Set sourceDocument = ActiveDocument
    editor.BeginEdit modeName, sourceDocument
    operation = "Menampilkan DirectorySettings"
    editor.Show vbModal
    Exit Sub
OpenFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    If Not editor Is Nothing Then Unload editor
    On Error GoTo 0
    MsgBox operation & vbCrLf & "Error " & CStr(errorNumber) & ": " & errorDescription, vbExclamation, "AutoSaveNCreate"
End Sub

Private Function SyncActiveDocument() As Boolean
    Dim doc As Document
    Dim detector As SNCModeDetector
    Dim modeName As String
    If Application.Documents.Count = 0 Then
        Set pModeDocument = Nothing
        RestoreDefaultMode
        Exit Function
    End If
    Set doc = ActiveDocument
    If Not pModeDocument Is Nothing Then
        If pModeDocument Is doc Then Exit Function
    End If
    Set detector = New SNCModeDetector
    modeName = detector.DetectMode(doc.FileName)
    RestoreDefaultMode
    Select Case modeName
        Case "DieA": ClearModeOptions: optDieA.Value = True
        Case "HiDie": ClearModeOptions: optHiDie.Value = True
        Case "KissA": ClearModeOptions: optKissA.Value = True
        Case "HiKiss": ClearModeOptions: optHiKiss.Value = True
    End Select
    Set pModeDocument = doc
    SyncActiveDocument = True
End Function

Private Sub ClearModeOptions()
    optDieA.Value = False
    optHiDie.Value = False
    optKissA.Value = False
    optHiKiss.Value = False
End Sub

Private Sub RestoreDefaultMode()
    If IsEmpty(pDefaultModes) Then Exit Sub
    ClearModeOptions
    optDieA.Value = pDefaultModes(0)
    optHiDie.Value = pDefaultModes(1)
    optKissA.Value = pDefaultModes(2)
    optHiKiss.Value = pDefaultModes(3)
End Sub
