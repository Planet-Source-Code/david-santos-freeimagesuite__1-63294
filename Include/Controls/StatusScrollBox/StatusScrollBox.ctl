VERSION 5.00
Begin VB.UserControl StatusScrollBox 
   ClientHeight    =   4230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5625
   ScaleHeight     =   4230
   ScaleWidth      =   5625
   Begin VB.PictureBox picIcons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      Picture         =   "StatusScrollBox.ctx":0000
      ScaleHeight     =   240
      ScaleWidth      =   1200
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.HScrollBar HScroll1 
      Enabled         =   0   'False
      Height          =   255
      LargeChange     =   10
      Left            =   0
      Max             =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3720
      Width           =   5175
   End
   Begin VB.VScrollBar VScrollBar 
      Enabled         =   0   'False
      Height          =   3735
      LargeChange     =   10
      Left            =   5160
      Max             =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picBG 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   3735
      Left            =   0
      ScaleHeight     =   3675
      ScaleWidth      =   5115
      TabIndex        =   0
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "StatusScrollBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim lastMaxLines As Long
Dim destDC As Long
Dim srcDC As Long
Dim clientHeight As Long
Dim clientWidth As Long
Dim mAutoRefresh As Boolean
Dim mAutoScroll As Boolean
Dim codeScrollChange As Boolean

Public Enum ICON_TYPE
    ICON_NOTICE = 0
    ICON_WARNING = 1
    ICON_ERROR = 2
End Enum

#If False Then
    Const ICON_NOTICE = 0
    Const ICON_WARNING = 1
    Const ICON_ERROR = 2
#End If

Dim mLine As Long
Dim mLines As Long
Dim asLines() As String
Dim asIcons() As ICON_TYPE

Dim maxLines As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

Dim WithEvents VScroll1 As CLongScroll
Attribute VScroll1.VB_VarHelpID = -1

Property Get AutoScroll() As Boolean
    AutoScroll = mAutoScroll
End Property

Property Let AutoScroll(v As Boolean)
    mAutoScroll = v
End Property

Property Get AutoRefresh() As Boolean
    AutoRefresh = mAutoRefresh
End Property

Property Let AutoRefresh(v As Boolean)
    mAutoRefresh = v
End Property

Private Sub HScroll1_Change()
    Refresh
End Sub

Private Sub HScroll1_Scroll()
    Refresh
End Sub

Private Sub UserControl_Initialize()
    Clear
    Set VScroll1 = New CLongScroll
    Set VScroll1.Client = VScrollBar
    mAutoRefresh = True
    mAutoScroll = True
End Sub

Public Sub Clear()
    maxLines = 1000
    lastMaxLines = 1000
    ReDim asLines(maxLines)
    ReDim asIcons(maxLines)
    mLine = 0
    mLines = 0
End Sub

Public Sub Refresh()
Dim scrollV As Long
Dim scrollH As Long
Dim iconTop As Long

    picBG.Cls
    
    scrollV = VScroll1.Value
    scrollH = HScroll1.Value
    
    destDC = picBG.hDC
    srcDC = picIcons.hDC
    
    For i = -scrollV To mLines - 1
        iconTop = (scrollV + i) * 16
        If iconTop >= 0 Then
            BitBlt destDC, -scrollH, iconTop, 16, 16, srcDC, CLng(asIcons(i)) * 16, 0, SRCCOPY
            TextOut destDC, -scrollH + 18, iconTop, asLines(i), Len(asLines(i))
        End If
        If iconTop > clientHeight Then Exit For
    Next
    picBG.Refresh
End Sub

Public Sub SetText(sText As String, Optional icontype As ICON_TYPE = ICON_NOTICE, Optional AutoAdvance As Boolean = True)
    If IsMissing(icontype) Then icontype = ICON_NOTICE
    
    If AutoAdvance Then
        asLines(mLines) = sText
        asIcons(mLines) = icontype
        mLine = mLines
        mLines = mLines + 1
    Else
        asLines(mLines - 1) = sText
        asIcons(mLines - 1) = icontype
    End If
    
    If mLines > maxLines Then
        maxLines = maxLines * 2
        ReDim Preserve asLines(maxLines)
        ReDim Preserve asIcons(maxLines)
    End If

    If mLines > (picBG.Height / (16 * 15)) Then
        VScroll1.Max = (picBG.Height / (16 * 15)) - mLines - 1
        
        VScrollBar.Enabled = True
        
        If mAutoScroll Then
            codeScrollChange = True
            VScroll1.Value = VScroll1.Max
            codeScrollChange = False
        End If
    Else
        VScrollBar.Enabled = False
    End If

    If mAutoRefresh Then Refresh
End Sub

Property Let CurrentLine(lLine As Long)
    mLine = lLine
End Property

Property Get CurrentLine() As Long
    CurrentLine = mLine
End Property

Property Get TotalLines() As Long
    TotalLines = mLines
End Property

Private Sub UserControl_Resize()
    If UserControl.Width - VScrollBar.Width < 0 Then Exit Sub
    If UserControl.Height - HScroll1.Height < 0 Then Exit Sub
    
    picBG.Width = UserControl.Width - VScrollBar.Width
    picBG.Height = UserControl.Height - HScroll1.Height
    
    VScrollBar.Height = picBG.Height
    HScroll1.Width = picBG.Width
    
    VScrollBar.Left = picBG.Width
    HScroll1.Top = picBG.Height

    clientHeight = picBG.Height \ Screen.TwipsPerPixelY
    clientWidth = picBG.Width \ Screen.TwipsPerPixelX

End Sub

Private Sub VScroll1_Change()
    If Not codeScrollChange Then
        If VScroll1.Value > VScroll1.Max Then
            mAutoScroll = False
        Else
            mAutoScroll = True
        End If
        Refresh
    End If
End Sub

Private Sub VScroll1_Scroll()
    Refresh
    If VScroll1.Value > VScroll1.Max Then mAutoScroll = False Else mAutoScroll = True
End Sub

Property Get Log() As String
Dim sBuffer As String
Dim buflen As Long
Dim sData As String
Dim lSize As Long

    buflen = 1024
    sBuffer = Space(buflen)
    bufpos = 1
    
    For i = 0 To mLines - 1
        sData = asLines(i) & vbCrLf
        lSize = Len(sData)
        Mid$(sBuffer, bufpos, lSize) = sData
        bufpos = bufpos + lSize
        If bufpos > buflen Then
            sBuffer = sBuffer & Space(buflen)
            buflen = Len(sBuffer)
        End If
    Next

    Log = Left(sBuffer, bufpos)
    sBuffer = ""
End Property
