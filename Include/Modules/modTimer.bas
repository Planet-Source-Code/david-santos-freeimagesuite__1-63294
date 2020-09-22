Attribute VB_Name = "modTimer"
Option Explicit


'=============================================================================================================
'
' modTimer Module
' ---------------
'
' Created By  : Kevin Wilson
'               http://www.TheVBZone.com   ( The VB Zone )
'               http://www.TheVBZone.net   ( The VB Zone .net )
'
' Last Update : April 01, 2000
'
' VB Versions : 5.0 / 6.0
'
' Requires    : cTimer.cls  ( Main Class Module )
'
' NOTE        : This module uses the Windows multimedia timer APIs to do a type of subclassing that requires
'               that you terminate this class module before exiting your program.  If you compile your program
'               and run it, you'll not have any problems and don't have to worry about this.  However, if
'               you're running in debug mode and click the STOP button, the class never gets properly terminated.
'               This will not crash your program or cause any problems directly, but to avoid possible problems
'               with timer events being left open, please close all the open forms in your project to end a
'               debug run instead of clicking the STOP button in Visual Basic.  This will process the class's
'               terminate event.
'
' Description : This class module was designed to take the place of VB's default Timer control.  Unlike the VB
'               standard Timer control, this class module does not require a form to run, allows for more
'               control of the timer,and is accurate down to 1 millisecond on most systems. ( the Min property
'               will return what your system's smallest possible time measurement is )
'
' SEE ALSO    : cTimer_NoSC.cls
'               ( This version of the Timer class module is the subclassing version.  cTimer_NoSC.cls is the
'               No Subclassing version )
'
' Example Use :
'
'  Private Timer1 As New cTimer
'
'  Private Sub Form_Load()
'    Set Timer1 = New cTimer
'    Timer1.Interval = 2000
'    Timer1.Enabled = True
'  End Sub
'
'  Private Sub Form_Click()
'    Debug.Print Timer1.TimeElapsed
'  End Sub
'
'  * IMPORTANT - PUT TIMER RELATED FUNCTIONALITY IN THE TimeProc FUNCTION
'
'=============================================================================================================
'
' LEGAL:
'
' You are free to use this code as long as you keep the above heading information intact and unchanged. Credit
' given where credit is due.  Also, it is not required, but it would be appreciated if you would mention
' somewhere in your compiled program that that your program makes use of code written and distributed by
' Kevin Wilson (www.TheVBZone.com).  Feel free to link to this code via your web site or articles.
'
' You may NOT take this code and pass it off as your own.  You may NOT distribute this code on your own server
' or web site.  You may NOT take code created by Kevin Wilson (www.TheVBZone.com) and use it to create products,
' utilities, or applications that directly compete with products, utilities, and applications created by Kevin
' Wilson, TheVBZone.com, or Wilson Media.  You may NOT take this code and sell it for profit without first
' obtaining the written consent of the author Kevin Wilson.
'
' These conditions are subject to change at the discretion of the owner Kevin Wilson at any time without
' warning or notice.  Copyright© by Kevin Wilson.  All rights reserved.
'
'=============================================================================================================


' Public Types
Public Type TIMECAPS
  wPeriodMin As Long
  wPeriodMax As Long
End Type

' Public Windows API Declarations
Public Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Public Declare Function timeEndPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Public Declare Function timeGetDevCaps Lib "winmm.dll" (lpTimeCaps As TIMECAPS, ByVal uSize As Long) As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function timeSetEvent Lib "winmm.dll" (ByVal uDelay As Long, ByVal uResolution As Long, ByVal lpFunction As Long, ByVal dwUser As Long, ByVal uFlags As Long) As Long
Public Declare Function timeKillEvent Lib "winmm.dll" (ByVal uID As Long) As Long


' PUT CODE FOR TIMER RELATED FUNCTIONALITY HERE!
Public Sub TimeProc(ByVal uID As Long, ByVal uMsg As Long, ByVal dwUser As Long, ByVal dw1 As Long, ByVal dw2 As Long)
Dim i As Long

    mFileList.Path = mSource
    mFileList.Filter = mFilter
    mFileList.Scan
    mFileList.Sorted = True
    
    For i = 0 To mFileList.FileCount - 1
        If Dir(mDest & "\" & mFileList.File(i)) = "" Then
            FileCopyProgress mSource & "\" & mFileList.File(i), mDest & "\" & mFileList.File(i), True
        End If
    Next
End Sub

