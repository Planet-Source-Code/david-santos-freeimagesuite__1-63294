Attribute VB_Name = "modBrowse"
Private Type BROWSEINFO
  hOwner As Long
  pidlRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfnCallback As Long
  lParam As Long
  iImage As Long
End Type

Dim m_CurrentDirectory As String

Private Const BIF_STATUSTEXT = &H4&
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BFFM_SETSELECTION = (WM_USER + 102)

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Public Function BrowseFolder(szDialogTitle As String, Optional hOwner As Long = 0, Optional StartFrom As String = "") As String
  Dim X As Long, bi As BROWSEINFO, dwIList As Long
  Dim szPath As String, wPos As Integer
  
    m_CurrentDirectory = StartFrom
  
    With bi
        .hOwner = hOwner
        .lpszTitle = szDialogTitle
        .ulFlags = BIF_RETURNONLYFSDIRS
        '.lpfnCallback = GetAddressofFunction(AddressOf BrowseCallbackProc)
    End With
    
    dwIList = SHBrowseForFolder(bi)
    szPath = Space$(512)
    X = SHGetPathFromIDList(ByVal dwIList, ByVal szPath)
    
    If X Then
        wPos = InStr(szPath, Chr(0))
        BrowseFolder = Left$(szPath, wPos - 1)
    Else
        BrowseFolder = vbNullString
    End If

End Function


Private Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
  
  Dim lpIDList As Long
  Dim ret As Long
  Dim sBuffer As String
  
  On Error Resume Next  'Sugested by MS to prevent an error from
                        'propagating back into the calling process.
     
  Select Case uMsg
  
    Case BFFM_INITIALIZED
      
      If Len(m_CurrentDirectory) > 0 Then
          Call SendMessage(hWnd, BFFM_SETSELECTION, 1, m_CurrentDirectory)
      End If
      
    Case BFFM_SELCHANGED
      sBuffer = Space(MAX_PATH)
      
      ret = SHGetPathFromIDList(lp, sBuffer)
      If ret = 1 Then
        Call SendMessage(hWnd, BFFM_SETSTATUSTEXT, 0, sBuffer)
      End If
      
  End Select
  
  BrowseCallbackProc = 0
  
End Function

' This function allows you to assign a function pointer to a vaiable.
Private Function GetAddressofFunction(add As Long) As Long
  GetAddressofFunction = add
End Function

