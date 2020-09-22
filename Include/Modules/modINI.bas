Attribute VB_Name = "modINI"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

Public Function ReadINI(INIFile As String, section As String, Value As String, default As String) As String
    Dim tempstr As String
    Dim slength As Long                      ' receives length of the returned string
    tempstr = Space(255)                     ' provide enough room for the function to put the value into the buffer
    slength = GetPrivateProfileString(section, Value, default, tempstr, 255, INIFile)
    ReadINI = Left(tempstr, slength)     ' extract the returned string from the buffer
    If ReadINI = "" Then ReadINI = default
End Function

Public Sub WriteINI(INIFile As String, section As String, keyname As String, Value As String)
    Dim retval As Long
    retval = WritePrivateProfileString(section, keyname, Value, INIFile)
End Sub

Public Function ReadINISection(INIFile As String, section As String) As String()
    Dim tempstr As String
    Dim slength As Long                      ' receives length of the returned string
    tempstr = Space(255)                     ' provide enough room for the function to put the value into the buffer
    slength = GetPrivateProfileSection(section, tempstr, 255, INIFile)
    ReadINISection = Split(Left(tempstr, slength), Chr(0))    ' extract the returned string from the buffer
End Function


