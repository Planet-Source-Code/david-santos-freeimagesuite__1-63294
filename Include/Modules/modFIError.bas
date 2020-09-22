Attribute VB_Name = "modFIError"
Public Sub FreeImageError(format As FREE_IMAGE_FORMAT, message As String)
    Debug.Print message
End Sub

Public Function GetFarPointer(address As Long) As Long
    GetFarPointer = address
End Function

