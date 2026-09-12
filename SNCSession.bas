Option Explicit

' Standard Module (Name) = SNCSession.
' Bertahan saat UF ditutup/dibuka selama proyek VBA belum di-reset.
Private pDirectorySession As SNCDirectorySession

Public Function GetSNCDirectorySession() As SNCDirectorySession
    If pDirectorySession Is Nothing Then Set pDirectorySession = New SNCDirectorySession
    Set GetSNCDirectorySession = pDirectorySession
End Function
