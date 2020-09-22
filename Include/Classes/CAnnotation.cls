VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAnnotation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Implements IOverlay

Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

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
Dim mHitpoint As HitTestEnum

Property Get SelectedColor() As Long
    SelectedColor = mSelectedColor
End Property

Property Let SelectedColor(v As Long)
    mSelectedColor = v
End Property

Property Get FillColor() As Long
    FillColor = lFillColor
End Property

Property Let FillColor(v As Long)
    lFillColor = v
End Property

Property Get Filled() As Boolean
    Filled = bFilled
End Property

Property Let Filled(v As Boolean)
    bFilled = v
End Property

Property Get LineWidth() As Long
    LineWidth = mLineWidth
End Property

Property Let LineWidth(v As Long)
    mLineWidth = v
End Property

Property Get Rotation() As Long
    Rotation = mRotation
End Property

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
    lFillColor = RGB(255, 255, 0)
    bFilled = True
    Set coords = New CRect
End Sub

Private Sub Class_Terminate()
    Debug.Print "Destroying annotation"
End Sub

Private Property Get IOverlay_AnnotationType() As ANNOTATION_TYPE
    If bFilled Then
        IOverlay_AnnotationType = fieFilledRect
    Else
        IOverlay_AnnotationType = fieHollowRect
    End If
End Property

Private Function IOverlay_GetRect() As CRect
    Set IOverlay_GetRect = GetRect
End Function

Private Property Let IOverlay_HitPoint(RHS As HitTestEnum)
    mHitpoint = RHS
End Property

Private Function IOverlay_HitTest(X As Variant, Y As Variant, Container As IContainer) As HitTestEnum
    Dim tRect As CRect
    
    Set tRect = Container.ImageToScreen(coords)
    
    ' EARLY OUT 1: Test if it's even inside the area
    If PtInRect(X, Y, tRect.Left, tRect.Top, tRect.Right, tRect.Bottom) Then
        
        If PtInRect(X, Y, tRect.Left + 6, tRect.Top + 6, tRect.Right - 6, tRect.Bottom - 6) Then
            ' EARLY OUT 2: Test for center first since this area is much bigger
            mHitpoint = HIT_CENTER
            
        ElseIf PtInRect(X, Y, tRect.Left, tRect.Top, tRect.Left + 5, tRect.Top + 5) Then
            mHitpoint = HIT_TOP + HIT_LEFT
        
        ElseIf PtInRect(X, Y, tRect.Left + 6, tRect.Top, tRect.Right - 6, tRect.Top + 5) Then
            mHitpoint = HIT_TOP
            
        ElseIf PtInRect(X, Y, tRect.Right - 5, tRect.Top, tRect.Right, tRect.Top + 5) Then
            mHitpoint = HIT_TOP + HIT_RIGHT
        
        ElseIf PtInRect(X, Y, tRect.Right - 5, tRect.Top + 6, tRect.Right, tRect.Bottom - 6) Then
            mHitpoint = HIT_RIGHT
        
        ElseIf PtInRect(X, Y, tRect.Right - 5, tRect.Bottom - 5, tRect.Right, tRect.Bottom) Then
            mHitpoint = HIT_BOTTOM + HIT_RIGHT
        
        ElseIf PtInRect(X, Y, tRect.Left + 6, tRect.Bottom - 5, tRect.Right - 6, tRect.Bottom) Then
            mHitpoint = HIT_BOTTOM
        
        ElseIf PtInRect(X, Y, tRect.Left, tRect.Bottom - 5, tRect.Left + 5, tRect.Bottom) Then
            mHitpoint = HIT_BOTTOM + HIT_LEFT
        
        ElseIf PtInRect(X, Y, tRect.Left, tRect.Top + 6, tRect.Left + 5, tRect.Bottom - 6) Then
            mHitpoint = HIT_LEFT
        
        End If
    
    Else
        mHitpoint = HIT_NONE
    
    End If

    'Return visible hittest to caller
    IOverlay_HitTest = mHitpoint

    ' rotate hittest internally
    Select Case Abs(Container.Rotation)
    Case 90
        Select Case mHitpoint
        Case HIT_TOP
            mHitpoint = HIT_RIGHT
            
        Case HIT_TOP + HIT_RIGHT
            mHitpoint = HIT_BOTTOM + HIT_RIGHT
            
        Case HIT_RIGHT
            mHitpoint = HIT_BOTTOM
        
        Case HIT_BOTTOM + HIT_RIGHT
            mHitpoint = HIT_BOTTOM + HIT_LEFT
        
        Case HIT_BOTTOM
            mHitpoint = HIT_LEFT
        
        Case HIT_BOTTOM + HIT_LEFT
            mHitpoint = HIT_TOP + HIT_LEFT
        
        Case HIT_LEFT
            mHitpoint = HIT_TOP
        
        Case HIT_TOP + HIT_LEFT
            mHitpoint = HIT_TOP + HIT_RIGHT
        
        End Select
    
    Case 180
        Select Case mHitpoint
        Case HIT_TOP
            mHitpoint = HIT_BOTTOM
            
        Case HIT_TOP + HIT_RIGHT
            mHitpoint = HIT_BOTTOM + HIT_LEFT
            
        Case HIT_RIGHT
            mHitpoint = HIT_LEFT
        
        Case HIT_BOTTOM + HIT_RIGHT
            mHitpoint = HIT_TOP + HIT_LEFT
        
        Case HIT_BOTTOM
            mHitpoint = HIT_TOP
        
        Case HIT_BOTTOM + HIT_LEFT
            mHitpoint = HIT_TOP + HIT_RIGHT
        
        Case HIT_LEFT
            mHitpoint = HIT_RIGHT
        
        Case HIT_TOP + HIT_LEFT
            mHitpoint = HIT_BOTTOM + HIT_RIGHT
        
        End Select
    
    
    Case 270
        Select Case mHitpoint
        Case HIT_TOP
            mHitpoint = HIT_LEFT
            
        Case HIT_TOP + HIT_RIGHT
            mHitpoint = HIT_TOP + HIT_LEFT
            
        Case HIT_RIGHT
            mHitpoint = HIT_TOP
        
        Case HIT_BOTTOM + HIT_RIGHT
            mHitpoint = HIT_TOP + HIT_RIGHT
        
        Case HIT_BOTTOM
            mHitpoint = HIT_RIGHT
        
        Case HIT_BOTTOM + HIT_LEFT
            mHitpoint = HIT_BOTTOM + HIT_RIGHT
        
        Case HIT_LEFT
            mHitpoint = HIT_BOTTOM
        
        Case HIT_TOP + HIT_LEFT
            mHitpoint = HIT_BOTTOM + HIT_LEFT
        
        End Select
    End Select

    'Rotate hitpoit


    Set tRect = Nothing

End Function

Private Sub IOverlay_Move(X, Y, Container As IContainer)
    Dim dx As Single
    Dim dy As Single

    Select Case Container.Rotation
    Case 0, 360
        dx = X
        dy = Y

    Case 90, -270
        dx = -Y
        dy = X

    Case 180, -180
        dx = -X
        dy = -Y

    Case -90, 270
        dx = Y
        dy = -X

    End Select
    
    With coords
        If mHitpoint = HIT_CENTER Then
                ' better dragging behaviour at image edges
                If .Left + dx < 0 Then dx = -.Left
                If .Right + dx > Container.Width Then dx = Container.Width - .Right
                If .Top + dy < 0 Then dy = -.Top
                If .Bottom + dy > Container.Height Then dy = Container.Height - .Bottom
        
                .Left = .Left + dx
                .Top = .Top + dy
                .Right = .Right + dx
                .Bottom = .Bottom + dy
        Else
        
            ' better dragging behaviour at image edges
            If (mHitpoint And HIT_LEFT) = HIT_LEFT Then
                If .Left + dx < 0 Then dx = -.Left
                If .Left + dx > Container.Width Then dx = Container.Width - .Left
                If .Width + dx < 3 Then dx = 0
                .Left = .Left + dx
            End If
        
            If (mHitpoint And HIT_TOP) = HIT_TOP Then
                If .Top + dy < 0 Then dy = -.Top
                If .Top + dy > Container.Height Then dy = Container.Height - .Top
                If .Height + dy < 3 Then dy = 0
                .Top = .Top + dy
            End If
        
            If (mHitpoint And HIT_RIGHT) = HIT_RIGHT Then
                If .Right + dx < 0 Then dx = -.Right
                If .Right + dx > Container.Width Then dx = Container.Width - .Right
                If .Width + dx < 3 Then dx = 0
                .Right = .Right + dx
            End If
        
            If (mHitpoint And HIT_BOTTOM) = HIT_BOTTOM Then
                If .Bottom + dy < 0 Then dy = -.Bottom
                If .Bottom + dy > Container.Height Then dy = Container.Height - .Bottom
                If .Height + dy < 3 Then dy = 0
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
    Dim mPen As LOGPEN
    
    If mPage = Container.Page Then
        
        Set tRect = Container.ImageToScreen(coords)
                    
    
        If bFilled Then
            HighLight Container.hDC, tRect.Left, tRect.Top, tRect.Width, tRect.Height, lFillColor
        End If
        
        If bSelected Then
            mPen.lopnColor = mSelectedColor
            mPen.lopnStyle = PS_SOLID + PS_ENDCAP_ROUND
            mPen.lopnWidth.X = mLineWidth
        Else
            mPen.lopnColor = lLineColor
            mPen.lopnStyle = PS_SOLID + PS_ENDCAP_ROUND
            mPen.lopnWidth.X = mLineWidth
        End If
    
        hPen = CreatePenIndirect(mPen)
        oldPen = SelectObject(hDC, hPen)
        
        'Container.DrawMode = vbCopyPen
        SetROP2 Container.hDC, R2_COPYPEN
                    
        hOldObj = SelectObject(Container.hDC, hPen)
    
        MoveToEx Container.hDC, tRect.Left, tRect.Top, oldPoint
        LineTo Container.hDC, tRect.Right, tRect.Top
        LineTo Container.hDC, tRect.Right, tRect.Bottom
        LineTo Container.hDC, tRect.Left, tRect.Bottom
        LineTo Container.hDC, tRect.Left, tRect.Top

    
        ' Get rid of pen
        SelectObject Container.hDC, hOldObj
        DeleteObject hPen
    
    End If
End Sub

Private Function HighLight(ByVal hDC As Long, ByVal lLeft As Long, ByVal lTop As Long, _
                           ByVal lWidth As Long, ByVal lHeight As Long, ByVal lColor As Long) As Long
    Dim hBrush As Long
    Dim tRect As RECT
    Dim oldBrush As Long
    Dim myDC As New CMemoryDC

    myDC.Width = lWidth
    myDC.Height = lHeight

    tRect.Top = 0
    tRect.Left = 0
    tRect.Bottom = lWidth
    tRect.Right = lHeight

    hBrush = CreateSolidBrush(lColor)
    oldBrush = SelectObject(myDC.hDC, hBrush)
    'Rectangle myDC.hDC, 0, 0, 1, 1
    Rectangle myDC.hDC, 0, 0, lWidth, lHeight
    SelectObject myDC.hDC, oldBrush
    DeleteObject hBrush

    BitBlt hDC, lLeft, lTop, lWidth, lHeight, myDC.hDC, 0, 0, SRCAND

    Set myDC = Nothing
End Function


Private Sub IOverlay_Rotate(angle As Double)
'
End Sub

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



' to implement: fixed width fields e.g. OR NOT!
' L(20)
' R(20)
Public Function ToString(Optional ByVal StrFormat As String = "L,T,W,H,P,A") As String
Dim out As String
    out = StrFormat
    out = Replace(out, "L", coords.Left)
    out = Replace(out, "T", coords.Top)
    out = Replace(out, "R", coords.Right)
    out = Replace(out, "B", coords.Bottom)
    out = Replace(out, "W", coords.Right - coords.Left)
    out = Replace(out, "H", coords.Bottom - coords.Top)
    out = Replace(out, "P", mPage)
    out = Replace(out, "A", mRotation)
    out = Replace(out, "F", bFilled)
    ToString = out
End Function

Public Sub FromString(Fields As String, Optional ByVal StrFormat As String = "L,T,W,H,P,A")
    Err.Raise vbObjectError, "CAnnotation.FromString", "Not yet implemented"
End Sub

Public Function GetRect() As CRect
    Set GetRect = New CRect
    GetRect.CopyRect coords
End Function



Public Function GetCopy() As CAnnotation
    Set GetCopy = New CAnnotation
    GetCopy.Top = coords.Top
    GetCopy.Left = coords.Left
    GetCopy.Width = coords.Width
    GetCopy.Height = coords.Height
    GetCopy.LineWidth = mLineWidth
    GetCopy.SelectedColor = mSelectedColor
    GetCopy.Page = mPage
    GetCopy.Rotation = mRotation
   
   ' GetCopy.LineColor = lLineColor
    GetCopy.FillColor = lFillColor
    GetCopy.Filled = bFilled
End Function

