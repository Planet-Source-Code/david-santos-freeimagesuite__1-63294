Attribute VB_Name = "modFileSystem"
Option Explicit

Public Function GetFileName(ByVal Filepath As String, Optional RemoveExtenstion As Boolean = False, Optional Extension)
    If InStrRev(Filepath, "\") > 0 Then
        Filepath = Mid$(Filepath, InStrRev(Filepath, "\") + 1)
    End If

    If RemoveExtenstion Then
        If InStrRev(Filepath, ".") > 0 Then
            If Not IsMissing(Extension) Then
                Extension = Mid$(Filepath, InStrRev(Filepath, ".") + 1)
            End If

            Filepath = Left$(Filepath, InStrRev(Filepath, ".") - 1)
        End If
    End If

    GetFileName = Filepath
End Function

Public Function GetFilePath(ByVal Filepath As String)
    If InStrRev(Filepath, "\") > 0 Then
        Filepath = Left$(Filepath, InStrRev(Filepath, "\") - 1)
    Else
        Filepath = ""
    End If

    GetFilePath = Filepath
End Function

Public Function GetRelativePath(ByVal Filepath As String)
    Dim i As Long, p As Long
    Dim Skip As Long

    If InStr(1, Filepath, "\\") = 1 Then
        Skip = 4
    ElseIf InStr(1, Filepath, ":\") > 0 Then
        Skip = 1
    Else
        Err.Raise vbObjectError + 5, "GetRelativePath", "Not a full path"
    End If

    While i < Skip
        p = InStr(p + 1, Filepath, "\")
        i = i + 1
    Wend

    GetRelativePath = Mid(Filepath, p + 1)
End Function

Public Sub CreatePath(ByVal RootPath As String, ByVal Path As String)
    Dim asPath() As String
    Dim i As Long

    If Left(Path, 1) = "\" Then Path = Right(Path, Len(Path) - 1)
    If Right(Path, 1) = "\" Then Path = Left(Path, Len(Path) - 1)
    asPath = Split(Path, "\")
    Path = ""
    RootPath = UnQualifyPath(RootPath)
    For i = 0 To UBound(asPath)
        Path = Path & "\" & asPath(i)
        If Len(Dir(RootPath & Path, vbDirectory)) = 0 Then
            MkDir RootPath & Path
        End If
    Next
End Sub

Public Function QualifyPath(ByVal Path As String) As String
    If Len(Path) > 0 Then
        If Right(Path, 1) <> "\" Then Path = Path & "\"
    End If
    QualifyPath = Path
End Function

Public Function UnQualifyPath(ByVal Path As String) As String
    If Len(Path) > 0 Then
        If Right(Path, 1) = "\" Then Path = Left(Path, Len(Path) - 1)
    End If
    UnQualifyPath = Path
End Function

