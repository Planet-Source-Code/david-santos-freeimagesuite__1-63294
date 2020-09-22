Attribute VB_Name = "modFreeImgEdit"
Private Const LF_FACESIZE = 32
Public Const PI As Double = 3.14159265358979
Public Const TWOPI As Double = 6.28318530717958

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type Size
    cx As Long
    cy As Long
End Type

Public Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(1 To LF_FACESIZE) As Byte
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type LOGPEN
    lopnStyle As Long
    lopnWidth As POINTAPI
    lopnColor As Long
End Type

Public Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type

Public MyCounter As Long

Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function GetTextExtentExPoint Lib "gdi32" Alias "GetTextExtentExPointA" (ByVal hdc As Long, ByVal lpszStr As String, ByVal cchString As Long, ByVal nMaxExtent As Long, lpnFit As Long, alpDx As Any, lpSize As Size) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function CreatePenIndirect Lib "gdi32" (lpLogPen As LOGPEN) As Long
Public Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Public Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Any) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal h As Long, ByVal w As Long, ByVal e As Long, ByVal O As Long, ByVal w As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal c As Long, ByVal OP As Long, ByVal cp As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function StretchDIBits Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, ByVal lpBits As Long, ByVal lpBitsInfo As Long, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Public Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hMultiBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByVal lpBits As Long, ByVal lpBI As Long, ByVal wUsage As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long

'************************************************************************************
'*                                 API Declarations
'************************************************************************************

Public Const MAX_LENGTH = 512

' Pen Style Consts
Public Const PS_NULL = 5&
Public Const PS_SOLID = 0&
Public Const PS_STYLE_MASK = &HF&
Public Const PS_TYPE_MASK = &HF0000
Public Const PS_JOIN_ROUND = &H0&
Public Const PS_JOIN_MITER = &H2000&
Public Const PS_JOIN_MASK = &HF000&
Public Const PS_JOIN_BEVEL = &H1000&
Public Const PS_INSIDEFRAME = 6&
Public Const PS_GEOMETRIC = &H10000
Public Const PS_ENDCAP_SQUARE = &H100&
Public Const PS_DOT = 2&
Public Const PS_ENDCAP_MASK = &HF00&
Public Const PS_ENDCAP_FLAT = &H200&
Public Const PS_ENDCAP_ROUND = &H0&
Public Const PS_DASH = 1&
Public Const PS_DASHDOT = 3&
Public Const PS_DASHDOTDOT = 4&
Public Const PS_COSMETIC = &H0&
Public Const PS_ALTERNATE = 8&

' Source Blit Consts
Public Const SRCCOPY = &HCC0020
Public Const SRCERASE = &H440328
Public Const SRCINVERT = &H660046
Public Const SRCPAINT = &HEE0086
Public Const SRCAND = &H8800C6

Public Const TRANSPARENT = 1

Public Const R2_BLACK = 1       '   0
Public Const R2_COPYPEN = 13    '  P
Public Const R2_LAST = 16
Public Const R2_MASKNOTPEN = 3  '  DPna
Public Const R2_MASKPEN = 9     '  DPa
Public Const R2_MASKPENNOT = 5  '  PDna
Public Const R2_MERGENOTPEN = 12        '  DPno
Public Const R2_MERGEPEN = 15   '  DPo
Public Const R2_MERGEPENNOT = 14        '  PDno
Public Const R2_NOT = 6 '  Dn
Public Const R2_NOP = 11        '  D
Public Const R2_NOTCOPYPEN = 4  '  PN
Public Const R2_NOTMASKPEN = 8  '  DPan
Public Const R2_NOTMERGEPEN = 2 '  DPon
Public Const R2_NOTXORPEN = 10  '  DPxn
Public Const R2_WHITE = 16      '   1
Public Const R2_XORPEN = 7      '  DPx


' Font Weight Consts
Public Const FW_BOLD = 700
Public Const FW_DONTCARE = 0
Public Const FW_EXTRABOLD = 800
Public Const FW_EXTRALIGHT = 200
Public Const FW_HEAVY = 900
Public Const FW_LIGHT = 300
Public Const FW_MEDIUM = 500
Public Const FW_NORMAL = 400
Public Const FW_SEMIBOLD = 600
Public Const FW_THIN = 100

'Private Const DT_BOTTOM = &H8
'Private Const DT_CALCRECT = &H400
'Private Const DT_CENTER = &H1
'Private Const DT_CHARSTREAM = 4          '  Character-stream, PLP
'Private Const DT_DISPFILE = 6            '  Display-file
'Private Const DT_EXPANDTABS = &H40
'Private Const DT_EXTERNALLEADING = &H200
'Private Const DT_INTERNAL = &H1000
'Private Const DT_LEFT = &H0
'Private Const DT_METAFILE = 5            '  Metafile, VDM
'Private Const DT_NOCLIP = &H100
'Private Const DT_NOPREFIX = &H800
'Private Const DT_PLOTTER = 0             '  Vector plotter
'Private Const DT_RASCAMERA = 3           '  Raster camera
'Private Const DT_RASDISPLAY = 1          '  Raster display
'Private Const DT_RASPRINTER = 2          '  Raster printer
'Private Const DT_RIGHT = &H2
'Private Const DT_SINGLELINE = &H20
'Private Const DT_TABSTOP = &H80
'Private Const DT_TOP = &H0
'Private Const DT_WORDBREAK = &H10
'Private Const DT_VCENTER = &H4
'
'Private Const ETO_CLIPPED = 4
'Private Const ETO_GRAYED = 1
'Private Const ETO_OPAQUE = 2

'Private Declare Function SetTextAlign Lib "gdi32" (ByVal hdc As Long, ByVal wFlags As Long) As Long
'Private Const TA_BASELINE = 24
'Private Const TA_BOTTOM = 8
'Private Const TA_CENTER = 6
'Private Const TA_LEFT = 0
'Private Const TA_NOUPDATECP = 0
'Private Const TA_RIGHT = 2
'Private Const TA_TOP = 0
'Private Const TA_UPDATECP = 1

'Private Type BITMAP
'    bmType As Long
'    bmWidth As Long
'    bmHeight As Long
'    bmWidthBytes As Long
'    bmPlanes As Integer
'    bmBitsPixel As Integer
'    bmBits As Long
'End Type


'Public Function Atan2(y As Double, x As Double) As Double
'    If x > 0 Then
'        Atan2 = Atn(y / x)
'    ElseIf x < 0 Then
'        Atan2 = Sgn(y) * (PI - Atn(Abs(y / x)))
'    ElseIf y = 0 Then
'        Atan2 = 0
'    Else
'        Atan2 = Sgn(y) * PI / 2
'    End If
'End Function

Public Function PtInRect(tx, ty, x1, y1, x2, y2) As Boolean
    PtInRect = (tx >= x1 And tx <= x2 And ty >= y1 And ty <= y2)
End Function

Public Function Min(X, Y)
    Min = IIf(X < Y, X, Y)
End Function

Public Function Max(X, Y)
    Max = IIf(X > Y, X, Y)
End Function

Public Function InsidePolygon(X, Y, polygon() As POINTAPI) As Boolean
Dim counter As Integer
Dim i As Long
Dim xinters As Double
Dim p As POINTAPI, p1 As POINTAPI, p2 As POINTAPI
    
    p.X = X
    p.Y = Y

    p1 = polygon(0)
    For i = 1 To UBound(polygon)
        p2 = polygon(i)
        If p.Y > Min(p1.Y, p2.Y) Then
            If p.Y <= Max(p1.Y, p2.Y) Then
                If p.X <= Max(p1.X, p2.X) Then
                    If p1.Y <> p2.Y Then
                        xinters = (p.Y - p1.Y) * (p2.X - p1.X) / (p2.Y - p1.Y) + p1.X
                        If (p1.X = p2.X) Or (p.X <= xinters) Then
                            counter = counter + 1
                        End If
                    End If
                End If
            End If
        End If
        p1 = p2
    Next
    
    If counter Mod 2 = 0 Then
        InsidePolygon = False
    Else
        InsidePolygon = True
    End If
End Function



