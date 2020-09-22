Attribute VB_Name = "modCopyFile"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2005 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Const MAXDWORD As Long = &HFFFFFFFF
Public Const MAX_PATH As Long = 260
Public Const INVALID_HANDLE_VALUE As Long = -1
Public Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10

'Define possible return codes from the CopyFileEx callback routine
Public Const PROGRESS_CONTINUE As Long = 0
Public Const PROGRESS_CANCEL As Long = 1
Public Const PROGRESS_STOP As Long = 2
Public Const PROGRESS_QUIET As Long = 3

'CopyFileEx callback routine state change values
Public Const CALLBACK_CHUNK_FINISHED As Long = &H0
Public Const CALLBACK_STREAM_SWITCH As Long = &H1

'CopyFileEx option flags
Public Const COPY_FILE_FAIL_IF_EXISTS As Long = &H1
Public Const COPY_FILE_RESTARTABLE As Long = &H2
Public Const COPY_FILE_OPEN_SOURCE_FOR_WRITE As Long = &H4

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData _
                                                                                                         As WIN32_FIND_DATA) As Long

Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As _
                                                                                                    WIN32_FIND_DATA) As Long

Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Public Declare Function CompareFileTime Lib "kernel32" (lpFileTime1 As FILETIME, lpFileTime2 As FILETIME) As Long

Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal _
                                                                                                       lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Public Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, _
                                                                                 lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long

Public Declare Function CopyFileEx Lib "kernel32" Alias "CopyFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal lpProgressRoutine As Long, lpData As Any, pbCancel As Long, ByVal dwCopyFlags As Long) As Long

Public bCancel As Boolean
Public FinishedCopying As Boolean

Public Function FARPROC(ByVal pfn As Long) As Long
'A dummy procedure that receives and returns
'the value of the AddressOf operator.

'Obtain and set the address of the callback
'This workaround is needed as you can't assign
'AddressOf directly to a member of a user-
'defined type, but you can assign it to another
'long and use that (as returned here)

    FARPROC = pfn

End Function


Public Function CopyProgressCallback(ByVal TotalFileSize As Currency, ByVal TotalBytesTransferred As Currency, ByVal StreamSize As Currency, ByVal StreamBytesTransferred As Currency, ByVal dwStreamNumber As Long, ByVal dwCallbackReason As Long, ByVal hSourceFile As Long, ByVal hDestinationFile As Long, lpData As Long) As Long
    Dim hDC As Long

    Select Case dwCallbackReason
    Case CALLBACK_STREAM_SWITCH:

        'this value is passed whenever the
        'callback is initialized for each file.

        'frmInstallNow.ProgressBar1.Value = 0
        'frmInstallNow.ProgressBar1.Min = 0
        'frmInstallNow.ProgressBar1.Max = (TotalFileSize * 10000)

        'frmInstallNow.ProgressBar1.Refresh

        CopyProgressCallback = PROGRESS_CONTINUE

    Case CALLBACK_CHUNK_FINISHED

        'called when a block has been copied
        If Not bCancel Then
            With frmCopier
                .Parent.Progress .lFileCount, .lCurrPos, .sCurrFile, CLng((TotalBytesTransferred / TotalFileSize * 100))
            End With
        End If
        'optional. While the app is copying it
        'will not respond to input for canceling.
        DoEvents
        CopyProgressCallback = PROGRESS_CONTINUE

    End Select

End Function

