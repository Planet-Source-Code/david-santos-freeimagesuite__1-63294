Attribute VB_Name = "modString"
' Converts strings to proper case
' Any non-alphanumeric character will trigger the next occuring
' alphabetic character to be uppercased
Public Function Proper(ByVal strval As String) As String
    Dim r As Integer
    Dim ShouldBeUpperCase As Boolean

    Proper = ""

    If Len(Trim$(strval)) > 0 Then
        'Set to True for the first character
        ShouldBeUpperCase = True

        'lowercase it all!
        strval = LCase$(strval)

        'check each character
        For r = 1 To Len(strval)
            Select Case Mid$(strval, r, 1)
            Case " ", "-", ";"
                ShouldBeUpperCase = True
            
            Case Else
                If ShouldBeUpperCase Then
                    'replace the character
                    Mid$(strval, r, 1) = UCase(Mid$(strval, r, 1))
                    ShouldBeUpperCase = False
                End If
            End Select
        Next

        Proper = strval
    End If
End Function


'Escape all non-alphanumeric characters into XML-safe equivalents
Public Function Escape(ByVal sText As String) As String
    Dim i As Long
    Dim sOut As String
    Dim chrCode As Integer
    Dim lastPtr As Long
    
    sText = Trim(sText)

    If IsNull(sText) Or sText = "" Then
        Escape = sText
        Exit Function
    
    End If
    lastPtr = 1

    For i = 1 To Len(sText)
        chrCode = Asc(Mid$(sText, i, 1))
        Select Case chrCode
        Case 48 To 57, 65 To 90, 97 To 122, 32
            'Entity  ISO Character #
            ' 0-9      48-57
            ' A-Z      65-90
            ' a-z      97-122
            ' Space    32
            ' do absolutely nothing
        Case Else
            sOut = sOut & Mid$(sText, lastPtr, i - lastPtr) & "&#" & chrCode & ";"
            lastPtr = i + 1
        End Select
    Next i
    
    sOut = sOut & Mid$(sText, lastPtr, i - lastPtr)
    
    Escape = sOut
End Function


