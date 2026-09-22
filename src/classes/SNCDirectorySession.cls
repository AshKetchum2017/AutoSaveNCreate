Option Explicit

' Class Module (Name) = SNCDirectorySession. Daftar bersama, tanpa registry.
' Daftar bertahan selama sesi VBA; pilihan aktif terikat identitas dokumen.
Private pEntries As Collection
Private pActiveDocument As Document
Private pActivePath As String

Private Sub Class_Initialize()
    Set pEntries = New Collection
    ClearSelection
End Sub

Public Function NormalizeDirectory(ByVal folderPath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    folderPath = Trim$(folderPath)
    If Not (Left$(folderPath, 2) = "\\" Or Mid$(folderPath, 2, 2) = ":\") Then
        Err.Raise 5, "SNCDirectorySession", "Masukkan path directory lengkap."
    End If
    If Not fso.FolderExists(folderPath) Then Err.Raise 5, "SNCDirectorySession", "Pilih directory tujuan yang sudah ada dan valid."
    NormalizeDirectory = fso.GetFolder(folderPath).Path
End Function

Public Function AddDirectory(ByVal folderPath As String, ByVal displayText As String) As Long
    Dim i As Long
    folderPath = NormalizeDirectory(folderPath)
    For i = 1 To pEntries.Count
        If StrComp(DirectoryAt(i), folderPath, vbTextCompare) = 0 Then
            AddDirectory = i
            Exit Function
        End If
    Next i
    pEntries.Add Array(folderPath, displayText)
    AddDirectory = pEntries.Count
End Function

Public Property Get Count() As Long
    Count = pEntries.Count
End Property

Public Function DirectoryAt(ByVal index As Long) As String
    Dim item As Variant
    item = pEntries(index)
    DirectoryAt = CStr(item(0))
End Function

Public Function DisplayAt(ByVal index As Long) As String
    Dim item As Variant
    item = pEntries(index)
    DisplayAt = CStr(item(1))
End Function

Public Sub SelectDirectory(ByVal doc As Document, ByVal index As Long)
    Dim folderPath As String
    If doc Is Nothing Then Err.Raise 5, "SNCDirectorySession", "Tidak ada dokumen aktif untuk DisposableDirectory."
    folderPath = NormalizeDirectory(DirectoryAt(index))
    Set pActiveDocument = doc
    pActivePath = folderPath
End Sub

Public Function GetDirectory(ByVal doc As Document) As String
    If doc Is Nothing Or pActiveDocument Is Nothing Then Exit Function
    If pActiveDocument Is doc Then GetDirectory = pActivePath
End Function

Public Sub ClearSelection()
    pActivePath = vbNullString
    Set pActiveDocument = Nothing
End Sub

Public Sub RemoveDirectory(ByVal index As Long)
    If StrComp(DirectoryAt(index), pActivePath, vbTextCompare) = 0 Then ClearSelection
    pEntries.Remove index
End Sub

Public Sub ClearAll()
    ClearSelection
    Set pEntries = New Collection
End Sub
