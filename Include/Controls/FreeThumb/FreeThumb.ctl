VERSION 5.00
Begin VB.UserControl FreeThumb 
   ClientHeight    =   4230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4590
   ScaleHeight     =   4230
   ScaleWidth      =   4590
   ToolboxBitmap   =   "FreeThumb.ctx":0000
   Begin VB.Timer scrollTimer 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   600
      Top             =   3360
   End
   Begin VB.Timer ThumbMaker 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   3360
   End
   Begin VB.PictureBox picDisplay 
      AutoRedraw      =   -1  'True
      Height          =   3975
      Left            =   0
      ScaleHeight     =   3915
      ScaleWidth      =   4275
      TabIndex        =   2
      Top             =   0
      Width           =   4335
   End
   Begin VB.HScrollBar HScroll1v 
      Height          =   255
      Left            =   0
      Max             =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3960
      Width           =   4335
   End
   Begin VB.VScrollBar VScroll1v 
      Height          =   3975
      Left            =   4320
      Max             =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "FreeThumb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim WithEvents HScroll1 As CLongScroll
Attribute HScroll1.VB_VarHelpID = -1
Dim WithEvents VScroll1 As CLongScroll
Attribute VScroll1.VB_VarHelpID = -1

Dim DCBuffer As New CMemoryDC

Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Dim UpdateList As New Collection

Dim vScrollPosition As Long
Dim vScrollEnd As Long
Dim hScrollPosition As Long
Dim hScrollEnd As Long
Dim VTimer As Boolean
Dim HTimer As Boolean

Dim scrollStart As Long
Dim scrollTarget As Long

Dim lSelected As Long
Dim mShowNavigator As Boolean
Dim busy As Boolean
Dim mThumbs As New CThumbs

Dim mSelected As CThumb
Dim mCSelection As Collection

Dim clientHeight As Long
Dim clientWidth As Long
Dim mThumbSize As Long

Dim mThumbHeight As Long
Dim mThumbWidth As Long

Public Event ItemClick(Item As CThumb)

Public Event NavigatorChange(ByVal Left As Long, ByVal Top As Long)

Public Event AfterAnnotationCreate(ByVal ThumbIndex As Long, ByVal Index As Long)
Public Event AnnotationCreate(ByVal ThumbIndex As Long, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, ByVal Rotation As Long, ByVal Page As Long)
Public Event BeforeAnnotationCreate(ByVal ThumbIndex As Long, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, _
                                        Cancel As Boolean, ByRef AnnotationType As ANNOTATION_TYPE, _
                                        ByRef FillColor As Long, ByRef LineColor As Long, ByRef Filled As Boolean)
                                        
Public Event AnnotationRemove(ByVal ThumbIndex As Long, ByVal Index As Long)
Public Event AnnotationClick(ByVal ThumbIndex As Long, ByVal Index As Long)
Public Event AnnotationChange(ByVal ThumbIndex As Long, ByVal Index As Long, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, Cancel As Boolean)

Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'local variable(s) to hold property value(s)
Private mvarMaxSize As Long 'local copy
Dim mThumbnailsPerRow As Long
Dim bDeferredUpdates As Boolean
Dim bDrawSelection As Boolean

Dim sx1 As Long
Dim sy1 As Long
Dim sx2 As Long
Dim sy2 As Long

Dim lastSelected As Long

Private mRubberBand As RubberBandMode

Property Let Rubberband(v As RubberBandMode)
    mRubberBand = v
    PropertyChanged "Rubberband"
End Property

Property Get Rubberband() As RubberBandMode
    Rubberband = mRubberBand
End Property

Public Function AddAbsoluteAnnotation(ByVal ThumbIndex As Long, ByVal lLeft As Long, ByVal lTop As Long, ByVal lWidth As Long, ByVal lHeight As Long, ByVal nRotation As Long, ByVal AnnotationType As ANNOTATION_TYPE, _
                                    ByVal FillColor As Long, ByVal LineColor As Long, ByVal Filled As Boolean, Optional Events As Boolean = True) As Long
    Dim lRect As New CRect
    Dim lCancel As Boolean
    
    Dim tempPage As Long
    Dim mTempWidth As Long
    Dim mTempHeight As Long
    
    Dim tempImage As Long
    Dim tAnnotationType As ANNOTATION_TYPE
    
    Dim thisThumb As CThumb
    Set thisThumb = mThumbs(ThumbIndex)
    
    If Events Then
        RaiseEvent BeforeAnnotationCreate(ThumbIndex, lLeft, lTop, lWidth, lHeight, lCancel, AnnotationType, FillColor, LineColor, Filled)
        If lCancel Then Exit Function
    End If
    
    lRect.CreateRect2 lTop, lLeft, lWidth, lHeight
    
    If lRect.Left > thisThumb.ThumbWidth Or lRect.Top > thisThumb.ThumbHeight Then
        Exit Function
    End If

    Dim newAnnotation As Object
    
    Select Case tAnnotationType
    Case fieFilledRect, fieHollowRect
        Dim newA As New CAnnotation
        
        With newA
            .Top = lRect.Top
            .Left = lRect.Left
            .Bottom = lRect.Bottom
            .Right = lRect.Right
            .Rotation = nRotation
            .Filled = Filled
            .FillColor = FillColor
            .Page = 1
        End With
    
        Set newAnnotation = newA
    
    Case fieLine
        Dim newL As New CLine
        With newL
            .Top = lRect.Top
            .Left = lRect.Left
            .Bottom = lRect.Bottom
            .Right = lRect.Right
            .Rotation = nRotation
            .Page = 1
        End With
    
        Set newAnnotation = newL
    
    Case fieText
        Dim newC As New CText
        With newC
            .Top = lRect.Top
            .Left = lRect.Left
            .Bottom = lRect.Bottom
            .Right = lRect.Right
            .Rotation = nRotation
            .Page = 1
        End With
    
        Set newAnnotation = newC
    
    End Select
    
    With newAnnotation
        If .Left < 0 Then .Left = 0
        If .Top < 0 Then .Top = 0
        
        If .Right < 0 Then .Right = 0
        If .Bottom < 0 Then .Bottom = 0
        
        If .Left > thisThumb.ThumbWidth Then .Left = thisThumb.ThumbWidth
        If .Top > thisThumb.ThumbHeight Then .Top = thisThumb.ThumbHeight
    
        If .Right > thisThumb.ThumbWidth Then .Right = thisThumb.ThumbWidth
        If .Bottom > thisThumb.ThumbHeight Then .Bottom = thisThumb.ThumbHeight
    End With
    
    thisThumb.Annotations.Add newAnnotation
    
    If Events Then RaiseEvent AfterAnnotationCreate(ThumbIndex, thisThumb.Annotations.Count)

    Set lRect = Nothing
    
    AddAbsoluteAnnotation = thisThumb.Annotations.Count

End Function

Property Let DrawSelection(v As Long)
    bDrawSelection = v
End Property

Property Get DrawSelection() As Long
    DrawSelection = bDrawSelection
End Property

Property Get FreeImageNames() As String
    FreeImageNames = mListOfFreeImage
End Property

Property Get AnnotationCount() As Long
    AnnotationCount = mThumbs(lSelected).AnnotationCount
End Property

'-----------------------------------
Property Get AnnotationLeft(ByVal Index As Long, ByVal AnnotationIndex As Long) As Long
Dim Left As Long
Dim Top As Long
Dim Width As Long
Dim Height As Long
            mThumbs(Index).GetAnnotation AnnotationIndex, Left, Top, Width, Height
            AnnotationLeft = Left
End Property


Property Get AnnotationTop(ByVal Index As Long, ByVal AnnotationIndex As Long) As Long
Dim Left As Long
Dim Top As Long
Dim Width As Long
Dim Height As Long
            mThumbs(Index).GetAnnotation AnnotationIndex, Left, Top, Width, Height
            AnnotationTop = Top
End Property

Property Get AnnotationWidth(ByVal Index As Long, ByVal AnnotationIndex As Long) As Long
Dim Left As Long
Dim Top As Long
Dim Width As Long
Dim Height As Long
            mThumbs(Index).GetAnnotation AnnotationIndex, Left, Top, Width, Height
            AnnotationWidth = Width
End Property

Property Get AnnotationHeight(ByVal Index As Long, ByVal AnnotationIndex As Long) As Long
Dim Left As Long
Dim Top As Long
Dim Width As Long
Dim Height As Long
            mThumbs(Index).GetAnnotation AnnotationIndex, Left, Top, Width, Height
            AnnotationHeight = Height
End Property

Public Sub GetAbsoluteAnnotation(ByVal Index As Long, ByVal AnnotationIndex As Long, Left As Long, Top As Long, Width As Long, Height As Long, Rotation As Long, Page As Long)
    mThumbs(Index).GetAnnotation AnnotationIndex, Left, Top, Width, Height
End Sub

Public Sub MoveAnnotation(ByVal Index As Long, ByVal AnnotationIndex As Long, ByVal DeltaX As Long, ByVal DeltaY As Long)
    Dim dx As Long
    Dim dy As Long

    If Index > -1 Then
            dx = DeltaX
            dy = DeltaY
            mThumbs(Index).MoveAnnotation AnnotationIndex, dx, dy
    End If
End Sub

Public Sub ClearAnnotations(ByVal Index As Long)
    mThumbs(Index).ClearAnnotations
End Sub

Property Get ThumbnailsPerRow() As Long
    ThumbnailsPerRow = mThumbnailsPerRow
End Property

Property Let ThumbnailsPerRow(v As Long)
    If v < 1 Then v = 1
    mThumbnailsPerRow = v
    UpdateScrollMax
End Property

Public Sub AddAnnotation(Index As Variant, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long)
    mThumbs(Index).AddAnnotation Left, Top, Width, Height, False
End Sub

Public Sub EditAnnotation(Index As Variant, ByVal AnnotationIndex As Long, ByVal HitPoint As HitTestEnum, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long)
    mThumbs(Index).EditAnnotation AnnotationIndex, HitPoint, Left, Top
End Sub

Public Sub RemoveAnnotation(Index As Variant, AnnotationIndex As Long)
    mThumbs(Index).RemoveAnnotation AnnotationIndex
End Sub

Property Get ShowNavigator() As Boolean
    ShowNavigator = mShowNavigator
End Property

Property Let ShowNavigator(v As Boolean)
    mShowNavigator = v
    Refresh
End Property

Public Sub SetNavigator(Left, Top, Width, Height)
    If Not mSelected Is Nothing Then
        mSelected.SetNavigator Left, Top, Width, Height
    End If
End Sub

Public Property Let MaxSize(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MaxSize = 5
    mvarMaxSize = vData
    mThumbSize = mvarMaxSize + 20
    UpdateScrollMax
End Property

Public Property Get MaxSize() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MaxSize
    MaxSize = mvarMaxSize
End Property

Public Property Let ThumbnailWidth(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MaxSize = 5
    mThumbWidth = vData
End Property

Public Property Get ThumbnailWidth() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MaxSize
    If lSelected > -1 Then
        mThumbWidth = mThumbs(lSelected).ThumbWidth
        ThumbnailWidth = mThumbWidth
    End If
End Property

Public Property Let ThumbnailHeight(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MaxSize = 5
    mThumbHeight = vData
End Property

Public Property Get ThumbnailHeight() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MaxSize
    If lSelected > -1 Then
        mThumbHeight = mThumbs(lSelected).ThumbHeight
        ThumbnailHeight = mThumbHeight
    End If
End Property

'-------------SCROLL VALUES ------
Property Let ScrollX(pos As Long)
    If pos > HScroll1.Max Then pos = HScroll1.Max
    If pos < 0 Then pos = 0
    'bUpdate = False
    HScroll1.Value = pos
    'bUpdate = True
End Property

Property Let ScrollY(pos As Long)
    If pos > VScroll1.Max Then pos = VScroll1.Max
    If pos < 0 Then pos = 0
    'bUpdate = False
    VScroll1.Value = pos
    'bUpdate = True
End Property

Property Get ScrollX() As Long
    ScrollX = HScroll1.Value
End Property

Property Get ScrollY() As Long
    ScrollY = VScroll1.Value
End Property

'-------------SCROLL VALUES ------

Public Sub Refresh()
    picDisplay.Cls
    DrawThumbNails
    
    If bDrawSelection Then
        Dim preColor As Long
        preColor = picDisplay.ForeColor
        picDisplay.ForeColor = 0
        picDisplay.DrawMode = vbInvert
        picDisplay.DrawStyle = vbDot
        
        Select Case mRubberBand
        Case RUBBERBAND_BOX
            picDisplay.Line (sx1 * 15, sy1 * 15)-(sx2 * 15, sy2 * 15), , B
        Case RUBBERBAND_LINE
            picDisplay.Line (sx1 * 15, sy1 * 15)-(sx2 * 15, sy2 * 15)
        End Select
        
        picDisplay.DrawStyle = vbSolid
        picDisplay.DrawMode = vbCopyPen
        picDisplay.ForeColor = preColor
    End If
    
    picDisplay.Refresh
End Sub

Property Get Thumbnails() As CThumbs
    Set Thumbnails = mThumbs
End Property

Private Sub HScroll1_Change()
    If Not bDeferredUpdates Then Refresh
End Sub

Private Sub HScroll1_Scroll()
    If Not bDeferredUpdates Then Refresh
End Sub

Private Sub MakeSelectedVisible()
    If lSelected < 1 Then lSelected = 1
    If lSelected > mThumbs.Count Then lSelected = mThumbs.Count
    EnsureVisible lSelected
    Refresh
End Sub

Private Sub picDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
    
    Select Case KeyCode
    Case vbKeyUp
        lSelected = lSelected - mThumbnailsPerRow
        MakeSelectedVisible
    
    Case vbKeyDown
        lSelected = lSelected + mThumbnailsPerRow
        MakeSelectedVisible
    
    Case vbKeyLeft
        lSelected = lSelected - 1
        MakeSelectedVisible
    
    Case vbKeyRight
        lSelected = lSelected + 1
        MakeSelectedVisible
    
    Case vbKeySpace
        Set mSelected = SelectThumb(lSelected, Shift)
        
        If Not mSelected Is Nothing Then
            RaiseEvent ItemClick(mSelected)
        End If
        
        Refresh
    End Select

End Sub

Private Sub picDisplay_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Function SelectThumb(ByVal Index As Long, ByVal Shift As Integer) As CThumb
    Dim i As Long
    Dim multiple As Boolean
    
    If Not ((Shift And vbCtrlMask) = vbCtrlMask) Then
        For i = 1 To mThumbs.Count
            mThumbs(i).Selected = False
            'RaiseEvent ItemDeselect(mThumbs(i))
        Next
    Else
        multiple = True
    End If
    
    If (Shift And vbShiftMask) = vbShiftMask Then
        If lastSelected = -1 Then lastSelected = 0
        
        For i = 1 To mThumbs.Count
            mThumbs(i).Selected = False
            'RaiseEvent ItemDeselect(mThumbs(i))
        Next
        
        T = lSelected - lastSelected
        If T = 0 Then T = 1
        s = Sgn(T)
        
        For i = lastSelected To Index - s Step s
            mThumbs(i).Selected = Not mThumbs(i).Selected
            'If mThumbs(i).Selected Then
            '    RaiseEvent ItemSelect(mThumbs(i))
            'End If
        Next
    Else
        lastSelected = Index
    End If
    
    If Index > -1 Then
        mThumbs(Index).Selected = Not mThumbs(Index).Selected
        Set SelectThumb = mThumbs(Index)
    Else
        Set SelectThumb = Nothing
    End If
    
    bDrawSelection = False
End Function


Private Sub picDisplay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    
    If Button = 1 Then
        lSelected = HitTest(X / 15, Y / 15)
        
        Set mSelected = SelectThumb(lSelected, Shift)
    
        If Not mSelected Is Nothing Then
            mSelected.MouseDown Button, Shift, X / 15, Y / 15
            RaiseEvent ItemClick(mSelected)
        End If
        
        'ItemClick shouldn't refresh?!?
        Refresh
    
    End If
End Sub

Private Sub picDisplay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    If Not mSelected Is Nothing Then
        mSelected.MouseMove Button, Shift, X / Screen.TwipsPerPixelX, Y / Screen.TwipsPerPixelY
        UserControl.MousePointer = mSelected.MousePointer
    End If
    Refresh
End Sub

Private Sub picDisplay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If Not mSelected Is Nothing And Not mShowNavigator Then
        mSelected.MouseUp Button, Shift, X / Screen.TwipsPerPixelX, Y / Screen.TwipsPerPixelY
    End If
    Refresh
End Sub

Public Function AddFile(Filename As String, Optional sKey As String) As CThumb
    Set Add = mThumbs.Add(Filename, sKey)
    Set Add.ParentControl = Me
    UpdateScrollMax
End Function

Public Sub Remove(vntIndexKey As Variant)
    'create a new object
    mThumbs.Remove vntIndexKey

    UpdateScrollMax
End Sub

Private Sub UpdateScrollMax()
Dim vMax As Long, hMax As Long
    
    vMax = (Int((mThumbs.Count - 1) / mThumbnailsPerRow) * mThumbSize) + mThumbSize - (VScroll1.Height / 15) + 20
    If vMax > 0 Then
        VScroll1.Max = vMax
        VScroll1.Enabled = True
    Else
        VScroll1.Max = 0
        VScroll1.Enabled = False
    End If
    
    hMax = (mThumbnailsPerRow * mThumbSize) - (HScroll1.Width / 15) + 10
    If hMax > 0 Then
        HScroll1.Max = hMax
        HScroll1.Enabled = True
    Else
        HScroll1.Max = 0
        HScroll1.Enabled = False
    End If

End Sub

Private Sub scrollTimer_Timer()
    If vScrollEnd < 0 Then vScrollEnd = 0
    If vScrollEnd > VScroll1.Max Then vScrollEnd = VScroll1.Max
    
    If hScrollEnd < 0 Then hScrollEnd = 0
    If hScrollEnd > HScroll1.Max Then hScrollEnd = HScroll1.Max

    bDeferredUpdates = True
    If Abs((vScrollEnd - vScrollPosition) / 3) > 1 Then
        vScrollPosition = vScrollPosition + (vScrollEnd - vScrollPosition) / 3
        If vScrollPosition > 0 And vScrollPosition < VScroll1.Max Then
            VScroll1.Value = vScrollPosition
        Else
            vScrollPosition = vScrollEnd
        End If
    Else
        If vScrollPosition > 0 Then
            VScroll1.Value = vScrollEnd
        End If
        VTimer = False
    End If

    If Abs((hScrollEnd - hScrollPosition) / 3) > 1 Then
        hScrollPosition = hScrollPosition + (hScrollEnd - hScrollPosition) / 3
        If hScrollPosition > 0 And hScrollPosition < HScroll1.Max Then
            HScroll1.Value = hScrollPosition
        Else
            hScrollPosition = hScrollEnd
        End If
    Else
        If hScrollPosition > 0 Then
            HScroll1.Value = hScrollEnd
        End If
        HTimer = False
    End If
    bDeferredUpdates = False
    Refresh
    
    If Not HTimer And Not VTimer Then scrollTimer.Enabled = False
End Sub

Private Sub ThumbMaker_Timer()
    If UpdateList.Count > 0 Then
        If Not UpdateList(1).HasThumbnail Then
            UpdateList(1).CreateThumbnail
        End If
        UpdateList.Remove 1
        Refresh
        If UpdateList.Count = 0 Then ThumbMaker.Enabled = False
    End If
End Sub

Private Sub UserControl_Initialize()
    lSelected = -1
    mvarMaxSize = 230
    mThumbnailsPerRow = 4
    mThumbSize = 250
    
    Set HScroll1 = New CLongScroll
    Set VScroll1 = New CLongScroll
    
    Set HScroll1.Client = HScroll1v
    Set VScroll1.Client = VScroll1v
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mRubberBand = PropBag.ReadProperty("Rubberband", RUBBERBAND_BOX)
End Sub

Private Sub UserControl_Resize()
    clientHeight = UserControl.Height - HScroll1.Height
    clientWidth = UserControl.Width - VScroll1.Width
    
    picDisplay.Height = clientHeight
    picDisplay.Width = clientWidth
    
    VScroll1.Left = clientWidth
    HScroll1.Top = clientHeight
    
    VScroll1.Height = clientHeight
    HScroll1.Width = clientWidth
    
    DCBuffer.Width = clientWidth / 15
    DCBuffer.Height = clientHeight / 15
End Sub

'Property Let Selected(v As Long)
'Dim mThumb As CThumb
'Dim I As Long
'    If v < 1 Or v > mThumbs.Count Then
'        Exit Property
'    End If
'
'    If lSelected >= 1 Then
'        Set mThumb = mThumbs(lSelected)
'        mThumb.Selected = False
'    End If
'
'    lSelected = v
'
'    Set mThumb = mThumbs(lSelected)
'    mThumb.Selected = True
'End Property

Property Get Selected() As CThumb
    Set Selected = mSelected
End Property

Public Function HitTest(X As Long, Y As Long) As Long
Dim mThumb As CThumb
Dim mTop As Long
Dim mLeft As Long
Dim mRight As Long
Dim mBottom As Long
Dim i As Long

    HitTest = -1
    For i = 1 To mThumbs.Count
        'Set mThumb = mThumbs(i)
        mTop = ((i - 1) \ mThumbnailsPerRow) * mThumbSize - VScroll1.Value
        mLeft = ((i - 1) Mod mThumbnailsPerRow) * mThumbSize - HScroll1.Value
        mRight = mLeft + mThumbSize
        mBottom = mTop + mThumbSize
        
        If (X > mLeft) And (X < mRight) And (Y > mTop) And (Y < mBottom) Then
            HitTest = i
            Exit For
        End If
    Next

End Function

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Rubberband", mRubberBand
End Sub

Private Sub VScroll1_Change()
    If Not bDeferredUpdates Then Refresh
End Sub

Private Sub VScroll1_Scroll()
    If Not bDeferredUpdates Then Refresh
End Sub

Public Sub DrawThumbNails()
Dim i As Long
Dim mThumb As CThumb
Dim lastpt As POINTAPI
Dim mTop As Long
Dim mLeft As Long
Dim mRight As Long
Dim mBottom As Long
Dim mBrush As LOGBRUSH
Dim hBrush As Long
Dim hPen As Long
Dim mPen As LOGPEN
Dim oldPen As Long
Dim tRect As RECT
Dim hDC As Long

    hDC = picDisplay.hDC
    
    For i = 1 To mThumbs.Count
        mLeft = ((i - 1) Mod mThumbnailsPerRow) * mThumbSize - HScroll1.Value
        mTop = ((i - 1) \ mThumbnailsPerRow) * mThumbSize - VScroll1.Value
        mRight = mLeft + mThumbSize
        mBottom = mTop + mThumbSize
        
        If BoxIsVisible(mLeft, mTop, mThumbSize, mThumbSize) Then
            Set mThumb = mThumbs(i)
            
            If Abs(mvarMaxSize - mThumb.MaxSize) > 0 Then
                mThumb.HasThumbnail = False
                mThumb.MaxSize = mvarMaxSize
            End If
            
            With tRect
                .Left = mLeft
                .Top = mTop
                .Right = mRight
                .Bottom = mBottom
            End With
            
            
            If mThumb.Selected Then
                mBrush.lbColor = RGB(128, 128, 128)
                hBrush = CreateBrushIndirect(mBrush)
                FillRect hDC, tRect, hBrush
                DeleteObject hBrush
            End If
                        
            If i = lSelected Then
                mPen.lopnColor = RGB(255, 255, 0)
            Else
                mPen.lopnColor = RGB(128, 0, 0)
            End If
            
            'mPen.lopnStyle = PS_SOLID + PS_ENDCAP_SQUARE
            mPen.lopnStyle = PS_SOLID + PS_ENDCAP_ROUND
            mPen.lopnWidth.X = 2
            
            hPen = CreatePenIndirect(mPen)
            oldPen = SelectObject(hDC, hPen)
            MoveToEx hDC, mLeft + 2, mTop + 2, lastpt
            LineTo hDC, mRight - 2, mTop + 2
            LineTo hDC, mRight - 2, mBottom - 2
            LineTo hDC, mLeft + 2, mBottom - 2
            LineTo hDC, mLeft + 2, mTop + 2
            SelectObject picDisplay.hDC, oldPen
            DeleteObject hPen
            
            If Not mThumb.HasThumbnail Then
                On Error Resume Next
                UpdateList.Add mThumb, mThumb.FileTitle
                If Err.Number = 0 And Not ThumbMaker.Enabled Then
                    ThumbMaker.Enabled = True
                End If
            End If
            
            Dim px As Long
            Dim py As Long
            
            px = mLeft + (mThumbSize - mThumb.ThumbWidth) / 2
            py = mTop + (mThumbSize - mThumb.ThumbHeight) / 2
            
            mThumb.DrawThumbnail hDC, px, py
            
            If i = lSelected And mShowNavigator Then
                mPen.lopnColor = RGB(0, 255, 0)
                mPen.lopnStyle = PS_SOLID + PS_ENDCAP_ROUND
                mPen.lopnWidth.X = 2
                        
                hPen = CreatePenIndirect(mPen)
                oldPen = SelectObject(hDC, hPen)
                
                With mThumb.Navigator
                    MoveToEx hDC, px + .Left, py + .Top, lastpt
                    LineTo hDC, px + .Right, py + .Top
                    LineTo hDC, px + .Right, py + .Bottom
                    LineTo hDC, px + .Left, py + .Bottom
                    LineTo hDC, px + .Left, py + .Top
                End With
                
                SelectObject hDC, oldPen
                DeleteObject hPen
            End If
            
            TextOut hDC, mLeft + 10, mBottom - 20, mThumb.FileTitle, Len(mThumb.FileTitle)
        
        End If
    Next
End Sub

Private Function BoxIsVisible(X, Y, w, h) As Boolean
    BoxIsVisible = True
    If X + w < 0 Or Y + h < 0 Or _
      X > (picDisplay.Width / 15) Or Y > (picDisplay.Height / 15) Then
        BoxIsVisible = False
    End If
End Function

Public Sub EnsureVisible(Index As Variant)
Static lastIndex As Long
Dim mTop As Long
Dim mLeft As Long

    mTop = ((Index - 1) \ mThumbnailsPerRow) * mThumbSize - VScroll1.Value
    
    If mTop < 0 Or mTop + mThumbSize > picDisplay.Height \ 15 Then
        vScrollPosition = VScroll1.Value
        ' smooth scrolling :)
    
        If Index > lastIndex Then
            vScrollEnd = -(picDisplay.Height \ 15 - mThumbSize - ((Index - 1) \ mThumbnailsPerRow) * mThumbSize)
        Else
            vScrollEnd = ((Index - 1) \ mThumbnailsPerRow) * mThumbSize
        End If
        
        
        lastIndex = Index
        
        VTimer = True
    End If

    mLeft = ((Index - 1) Mod mThumbnailsPerRow) * mThumbSize - HScroll1.Value

    If mLeft < 0 Or mLeft + mThumbSize > picDisplay.Width \ 15 Then
        hScrollPosition = HScroll1.Value
            
        If Index > lastIndex Then
            hScrollEnd = -(picDisplay.Width \ 15 - mThumbSize - ((Index - 1) Mod mThumbnailsPerRow) * mThumbSize)
        Else
            hScrollEnd = ((Index - 1) Mod mThumbnailsPerRow) * mThumbSize
        End If
        
        
        lastIndex = Index
    
        HTimer = True
    End If
        
    If HTimer Or VTimer Then scrollTimer.Enabled = True

End Sub

Public Sub SetSelection(x1, y1, x2, y2)
    sx1 = x1
    sy1 = y1
    sx2 = x2
    sy2 = y2
End Sub

