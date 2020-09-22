Attribute VB_Name = "modSysFolder"
Option Explicit

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" _
(ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
(ByVal pidl As Long, ByVal pszPath As String) As Long

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private Const MAX_PATH As Integer = 260

Public Enum CSIDL_CONSTANTS
    CSIDL_FLAG_CREATE = &H8000
    CSIDL_ADMINTOOLS = &H30
    CSIDL_ALTSTARTUP = &H1D
    CSIDL_APPDATA = &H1A
    CSIDL_BITBUCKET = &HA
    CSIDL_CDBURN_AREA = &H3B
    CSIDL_COMMON_ADMINTOOLS = &H2F
    CSIDL_COMMON_ALTSTARTUP = &H1E
    CSIDL_COMMON_APPDATA = &H23
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19
    CSIDL_COMMON_DOCUMENTS = &H2E
    CSIDL_COMMON_FAVORITES = &H1F
    CSIDL_COMMON_MUSIC = &H35
    CSIDL_COMMON_PICTURES = &H36
    CSIDL_COMMON_PROGRAMS = &H17
    CSIDL_COMMON_STARTMENU = &H16
    CSIDL_COMMON_STARTUP = &H18
    CSIDL_COMMON_TEMPLATES = &H2D
    CSIDL_COMMON_VIDEO = &H37
    CSIDL_CONTROLS = &H3
    CSIDL_COOKIES = &H21
    CSIDL_DESKTOP = &H0
    CSIDL_DESKTOPDIRECTORY = &H10
    CSIDL_DRIVES = &H11
    CSIDL_FAVORITES = &H6
    CSIDL_FONTS = &H14
    CSIDL_HISTORY = &H22
    CSIDL_INTERNET = &H1
    CSIDL_INTERNET_CACHE = &H20
    CSIDL_LOCAL_APPDATA = &H1C
    CSIDL_MYDOCUMENTS = &HC
    CSIDL_MYMUSIC = &HD
    CSIDL_MYPICTURES = &H27
    CSIDL_MYVIDEO = &HE
    CSIDL_NETHOOD = &H13
    CSIDL_NETWORK = &H12
    CSIDL_PERSONAL = &H5
    CSIDL_PRINTERS = &H4
    CSIDL_PRINTHOOD = &H1B
    CSIDL_PROFILE = &H28
    CSIDL_PROFILES = &H3E
    CSIDL_PROGRAM_FILES = &H26
    CSIDL_PROGRAM_FILES_COMMON = &H2B
    CSIDL_PROGRAMS = &H2
    CSIDL_RECENT = &H8
    CSIDL_SENDTO = &H9
    CSIDL_STARTMENU = &HB
    CSIDL_STARTUP = &H7
    CSIDL_SYSTEM = &H25
    CSIDL_TEMPLATES = &H15
    CSIDL_WINDOWS = &H24
End Enum

Public Function GetSpecialFolder(hWnd As Long, CSIDL As CSIDL_CONSTANTS) As String
    Dim sPath As String
    Dim IDL As ITEMIDLIST
    '
    ' Retrieve info about system folders such as the "Recent Documents" folder.
    ' Info is stored in the IDL structure.
    '
    GetSpecialFolder = ""
    If SHGetSpecialFolderLocation(hWnd, CSIDL, IDL) = 0 Then
        '
        ' Get the path from the ID list, and return the folder.
        sPath = Space$(MAX_PATH)
        If SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath) Then
            GetSpecialFolder = Left$(sPath, InStr(sPath, vbNullChar) - 1) & ""
        End If
    End If


End Function

