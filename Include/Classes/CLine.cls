VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements IOverlay

' here it's private
Private coords As CRect
Private mPage As Long
Private mRotation As Long
Private lLineColor As Long
Private bFilled As Boolean
Private lFillColor As Long
Private bFixed As Boolean
Private bSelected As Boolean
Private mLineWidth As Long
Private mSelectedColor As Long
Private mHitpoint As HitTestEnum

Private mText As String

Property Let Rotation(v As Long)
    mRotation = v
End Property

Property Let Page(v As Long)
    mPage = v
End Property

Property Let Top(v As Long)
    coords.Top = v
End Property

Property Let Left(v As Long)
    coords.Left = v
End Property

Property Let Bottom(v As Long)
    coords.Bottom = v
End Property

Property Let Right(v As Long)
    coords.Right = v
End Property

Property Let Width(v As Long)
    coords.Width = v
End Property

Property Let Height(v As Long)
    coords.Height = v
End Property

Property Get Top() As Long
    Top = coords.Top
End Property

Property Get Left() As Long
    Left = coords.Left
End Property

Property Get Bottom() As Long
    Bottom = coords.Bottom
End Property

Property Get Right() As Long
    Right = coords.Right
End Property

Property Get Width() As Long
    Width = coords.Width
End Property

Property Get Height() As Long
    Height = coords.Height
End Property

Private Sub Class_Initialize()
    mLineWidth = 2
    mSelectedColor = RGB(255, 0, 0)
    lLineColor = RGB(0, 0, 255)
    mHitpoint = HIT_NONE
    Set coords = New CRect
End Sub

Private Sub Class_Terminate()
    Debug.Print "Destroying annotation"
End Sub

Private Property Get IOverlay_AnnotationType() As ANNOTATION_TYPE
    IOverlay_AnnotationType = fieLine
End Property

Private Function IOverlay_GetRect() As CRect
    Set IOverlay_GetRect = GetRect
End Function

Private Function IOverlay_HitTest(x As Variant, y As Variant, Container As IContainer) As HitTestEnum
    Dim tRect As CRect
    
    Set tRect = Container.ImageToScreen(coords)
    
    If PtInRect(x, y, tRect.Left - 5, tRect.Top - 5, tRect.Left + 5, tRect.Top + 5) Then
        mHitpoint = HIT_START
    ElseIf PtInRect(x, y, tRect.Right - 5, tRect.Bottom - 5, tRect.Right + 5, tRect.Bottom + 5) Then
        mHitpoint = HIT_END
    ElseIf PtOnLine(x, y, tRect.Left, tRect.Top, tRect.Right, tRect.Bottom) Then
        mHitpoint = HIT_MOVE
    Else
        mHitpoint = HIT_NONE
    End If

    'Return visible hittest to caller
    IOverlay_HitTest = mHitpoint

    Set tRect = Nothing

End Function

Private Sub IOverlay_Move(x, y, Container As IContainer)
    Dim dx As Single
    Dim dy As Single

    Select Case Container.Rotation
    Case 0, 360
        dx = x
        dy = y

    Case 90, -270
        dx = -y
        dy = x

    Case 180, -180
        dx = -x
        dy = -y

    Case -90, 270
        dx = y
        dy = -x

    End Select
    
    With coords
        ' better dragging behaviour at image edges
        If mHitpoint = HIT_MOVE Then
            ' make sure both points don't go offscreen
            
            If .Left + dx < 0 Then dx = -.Left
            If .Right + dx < 0 Then dx = -.Right
            
            If .Left + dx > Container.Width Then dx = Container.Width - .Left
            If .Right + dx > Container.Width Then dx = Container.Width - .Right
            
            If .Top + dy < 0 Then dy = -.Top
            If .Bottom + dy < 0 Then dy = -.Bottom
            
            If .Top + dy > Container.Height Then dy = Container.Height - .Top
            If .Bottom + dy > Container.Height Then dy = Container.Height - .Bottom
        
    
            .Left = .Left + dx
            .Top = .Top + dy
            .Right = .Right + dx
            .Bottom = .Bottom + dy
        Else
        
            ' better dragging behaviour at image edges
            If (mHitpoint And HIT_START) = HIT_START Then
                ' make sure start point doesn't go offscreen
                
                If .Left + dx < 0 Then dx = -.Left
                If .Left + dx > Container.Width Then dx = Container.Width - .Left
                If .Top + dy < 0 Then dy = -.Top
                If .Top + dy > Container.Height Then dy = Container.Height - .Top
                
                .Left = .Left + dx
                .Top = .Top + dy
            End If
        
            If (mHitpoint And HIT_END) = HIT_END Then
                ' make sure end point doesn't go offscreen
                
                If .Right + dx < 0 Then dx = -.Right
                If .Right + dx > Container.Width Then dx = Container.Width - .Right
                If .Bottom + dy < 0 Then dy = -.Bottom
                If .Bottom + dy > Container.Height Then dy = Container.Height - .Bottom
                
                .Right = .Right + dx
                .Bottom = .Bottom + dy
            End If
        
        End If
    End With

End Sub

Private Sub IOverlay_Render(Container As IContainer)
    Dim oldPoint As POINTAPI
    Dim lCurBrush As Long
    Dim hPen As Long
    Dim hOldObj As Long
    Dim tRect As New CRect
    
    If mPage = Container.Page Then
        
        Set tRect = Container.ImageToScreen(coords)
    
        If bSelected Then
        '    'lCurBrush = hSelectedBrush
            
            hPen = CreatePen(PS_SOLID, mLineWidth, mSelectedColor)
            Container.DrawMode = vbInvert
            Container.ForeColor = RGB(0, 0, 0)
        
            MoveToEx Container.hDC, tRect.Left - 5, tRect.Top - 5, oldPoint
            LineTo Container.hDC, tRect.Left + 5, tRect.Top - 5
            LineTo Container.hDC, tRect.Left + 5, tRect.Top + 5
            LineTo Container.hDC, tRect.Left - 5, tRect.Top + 5
            LineTo Container.hDC, tRect.Left - 5, tRect.Top - 5
        
            MoveToEx Container.hDC, tRect.Right - 5, tRect.Bottom - 5, oldPoint
            LineTo Container.hDC, tRect.Right + 5, tRect.Bottom - 5
            LineTo Container.hDC, tRect.Right + 5, tRect.Bottom + 5
            LineTo Container.hDC, tRect.Right - 5, tRect.Bottom + 5
            LineTo Container.hDC, tRect.Right - 5, tRect.Bottom - 5
        
            Container.ForeColor = RGB(0, 0, 255)
        
        Else
            hPen = CreatePen(PS_SOLID, mLineWidth, lLineColor)
            Container.ForeColor = RGB(0, 0, 255)
        
        '    'lCurBrush = hBlackBrush
        End If
    
        Container.DrawMode = vbCopyPen
                    
        hOldObj = SelectObject(Container.hDC, hPen)
    
        MoveToEx Container.hDC, tRect.Left, tRect.Top, oldPoint
        LineTo Container.hDC, tRect.Right, tRect.Bottom
        
        ' Get rid of pen
        SelectObject Container.hDC, hOldObj
        DeleteObject hPen
    
    End If
End Sub

Private Sub IOverlay_Rotate(angle As Double)
'
End Sub

Private Function PtOnLine(tx, ty, x1, y1, x2, y2) As Boolean
Dim Slope As Single
Dim points(4) As POINTAPI

    'Slope = (y2 - y1) / (x2 - x1)
    
    If x2 >= x1 Then
        points(0).x = x1 - 5
        points(0).y = y1 + 5
        points(1).x = x1 + 5
        points(1).y = y1 - 5
        points(2).x = x2 + 5
        points(2).y = y2 - 5
        points(3).x = x2 - 5
        points(3).y = y2 + 5
        points(4).x = points(0).x
        points(4).y = points(0).y
    Else
        points(0).x = x1 - 5
        points(0).y = y1 - 5
        points(1).x = x1 + 5
        points(1).y = y1 + 5
        points(2).x = x2 + 5
        points(2).y = y2 + 5
        points(3).x = x2 - 5
        points(3).y = y2 - 5
        points(4).x = points(0).x
        points(4).y = points(0).y
    End If
    
    
    PtOnLine = InsidePolygon(tx, ty, points)
End Function


Private Property Let IOverlay_Top(v As Long)
    coords.Top = v
End Property

Private Property Let IOverlay_Left(v As Long)
    coords.Left = v
End Property

Private Property Let IOverlay_Bottom(v As Long)
    coords.Bottom = v
End Property

Private Property Let IOverlay_Right(v As Long)
    coords.Right = v
End Property

Private Property Let IOverlay_Width(v As Long)
    coords.Width = v
End Property

Private Property Let IOverlay_Height(v As Long)
    coords.Height = v
End Property

Private Property Let IOverlay_Selected(RHS As Boolean)
    bSelected = RHS
End Property

Private Property Let IOverlay_Rotation(RHS As Long)
    mRotation = RHS
End Property

Private Property Let IOverlay_Page(RHS As Long)
    mPage = RHS
End Property



Private Property Get IOverlay_Top() As Long
    IOverlay_Top = coords.Top
End Property

Private Property Get IOverlay_Left() As Long
    IOverlay_Left = coords.Left
End Property

Private Property Get IOverlay_Bottom() As Long
    IOverlay_Bottom = coords.Bottom
End Property

Private Property Get IOverlay_Right() As Long
    IOverlay_Right = coords.Right
End Property

Private Property Get IOverlay_Width() As Long
    IOverlay_Width = coords.Width
End Property

Private Property Get IOverlay_Height() As Long
    IOverlay_Height = coords.Height
End Property

Private Property Get IOverlay_Selected() As Boolean
    IOverlay_Selected = bSelected
End Property

Private Property Get IOverlay_Rotation() As Long
    IOverlay_Rotation = mRotation
End Property

Private Property Get IOverlay_Page() As Long
    IOverlay_Page = mPage
End Property


Private Property Let IOverlay_HitPoint(RHS As HitTestEnum)
    mHitpoint = RHS
End Property

' to implement: fixed width fields e.g. OR NOT!
' L(20)
' R(20)
Public Function ToString(Optional ByVal StrFormat As String = "L,T,W,H,P") As String
Dim out As String
    out = StrFormat
    out = Replace(out, "L", coords.Left)
    out = Replace(out, "T", coords.Top)
    out = Replace(out, "R", coords.Right)
    out = Replace(out, "B", coords.Bottom)
    out = Replace(out, "W", coords.Right - coords.Left)
    out = Replace(out, "H", coords.Bottom - coords.Top)
    out = Replace(out, "P", mPage)
    ToString = out
End Function

Public Sub FromString(Fields As String, Optional ByVal StrFormat As String = "L,T,W,H,P")
    Err.Raise vbObjectError, "CLine.FromString", "Not yet implemented"
End Sub

Public Function GetRect() As CRect
    Set GetRect = New CRect
    GetRect.CopyRect coords
End Function

Public Function GetStart() As CPoint
    Set GetStart = New CPoint
    GetStart.x = coords.Left
    GetStart.y = coords.Top
End Function

Public Function GetEnd() As CPoint
    Set GetEnd = New CPoint
    GetEnd.x = coords.Right
    GetEnd.y = coords.Bottom
End Function

Public Sub SetHitPoint(HitPoint As HitTestEnum)
    mHitpoint = HitPoint
End Sub

Public Property Get LineWidth() As Long
    LineWidth = mLineWidth
End Property

Public Property Let LineWidth(v As Long)
    mLineWidth = v
End Property

