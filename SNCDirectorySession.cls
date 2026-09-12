Option Explicit

' Class Module (Name) = SNCDirectorySession. Tidak menulis registry.
Private pEntries As Collection

Private Sub Class_Initialize()
    Set pEntries = New Collection
End Sub

Public Sub SetDirectory(ByVal doc As Document, ByVal modeName As String, ByVal folderPath As String)
    Dim fso As Object
    Dim entry As Collection
    If doc Is Nothing Then Err.Raise 5, "SNCDirectorySession", "Tidak ada dokumen aktif untuk DisposableSave."
    Select Case modeName
        Case "DieA", "HiDie", "KissA", "HiKiss"
        Case Else
            Err.Raise 5, "SNCDirectorySession", "Mode DisposableSave tidak dikenal."
    End Select
    Set fso = CreateObject("Scripting.FileSystemObject")
    folderPath = Trim$(folderPath)
    If Not fso.FolderExists(folderPath) Then Err.Raise 5, "SNCDirectorySession", "Pilih directory sekali pakai yang valid."
    Set entry = New Collection
    entry.Add doc, "Document"
    entry.Add modeName, "Mode"
    entry.Add folderPath, "Directory"
    ClearDirectory doc, modeName
    pEntries.Add entry
End Sub

Public Function GetDirectory(ByVal doc As Document, ByVal modeName As String) As String
    Dim i As Long
    Dim entry As Collection
    Dim source As Document
    If doc Is Nothing Then Exit Function
    For i = 1 To pEntries.Count
        Set entry = pEntries(i)
        Set source = entry("Document")
        If source Is doc Then
            If CStr(entry("Mode")) = modeName Then
                GetDirectory = CStr(entry("Directory"))
                Exit Function
            End If
        End If
    Next i
End Function

Public Sub ClearDirectory(ByVal doc As Document, ByVal modeName As String)
    Dim i As Long
    Dim entry As Collection
    Dim source As Document
    For i = pEntries.Count To 1 Step -1
        Set entry = pEntries(i)
        Set source = entry("Document")
        If source Is doc Then
            If CStr(entry("Mode")) = modeName Then pEntries.Remove i
        End If
    Next i
End Sub
