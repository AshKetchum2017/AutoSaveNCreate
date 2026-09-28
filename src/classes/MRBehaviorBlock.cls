Option Explicit

' Class Module (Name) = MRBehaviorBlock. One occurrence in the selected queue.
Public MacroID As String
Public SourceText As String
Public Instructions As Collection
Public HasCommands As Boolean

Private Sub Class_Initialize()
    Set Instructions = New Collection
End Sub
