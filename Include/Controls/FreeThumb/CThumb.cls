VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CThumb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Implements IContainer
Implements INavigator

Dim mNavigatorEvents As INavigatorEvents

Private Declare Function GetTickCount Lib "kernel32" () As Long

Dim T As Long

Dim hImage As Long
Dim sFilename As String
Dim sFileTitle As String
Dim mWidth As Long
Dim mHeight As Long
Dim mOrigWidth As Long
Dim mOrigHeight As Long
Dim sError As String
Dim lCurPage As Long
Dim mSelected As Boolean
Dim mSize As Long
Dim mHasThumbnail As Boolean
Dim mNavigatorSelected As Boolean

Dim mMousePointer As MousePointerConstants

Dim px As Long
Dim py As Long

Dim sX As Long
Dim sY As Long

'Dim mAnnotationCount As Long
Dim lSelected As Long
Dim mNavigator As CRect
Dim vZoom As Single

Dim WithEvents mCAnnotations As CAnnotations
Attribute mCAnnotations.VB_VarHelpID = -1
Dim WithEvents mCSelection As CAnnotations
Attribute mCSelection.VB_VarHelpID = -1

Dim mParentControl As FreeThumb

Dim mhDC As Long
Dim offsetX As Long
Dim offsetY As Long

Property Get MousePointer() As MousePointerConstants
    MousePointer = mMousePointer
End Property

Property Set ParentControl(v As Object)
    Set mParentControl = v
End Property

Property Get ParentControl() As Object
    Set ParentControl = mParentControl
End Property

Public Sub SetNavigator(Left, Top, Width, Height)
    mNavigator.Left = Left * vZoom
    mNavigator.Top = Top * vZoom
    mNavigator.Right = (Left + Width) * vZoom
    mNavigator.Bottom = (Top + Height) * vZoom
    'INavigator_Update
    mParentControl.Refresh
End Sub

Property Get Navigator() As CRect
    Set Navigator = mNavigator
End Property

Public Sub AddAnnotation(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, Optional Events As Boolean = True)
    If Width < 3 Or Height < 3 Then Exit Sub
    
    Dim newAnnotation As New CAnnotation
    With newAnnotation
        .Left = Left
        .Top = Top
        .Width = Width
        .Height = Height
        .Page = 1
        .Rotation = 0
    End With
    
    mCAnnotations.Add newAnnotation
    mCAnnotations.CallUpdate
    
End Sub

Property Get ThumbWidth() As Long
    ThumbWidth = mWidth
End Property

Property Get ThumbHeight() As Long
    ThumbHeight = mHeight
End Property

Property Let HasThumbnail(v As Boolean)
    mHasThumbnail = v
End Property

Property Get HasThumbnail() As Boolean
    HasThumbnail = mHasThumbnail
End Property

Property Let Selected(v As Boolean)
    mSelected = v
End Property

Property Get Selected() As Boolean
    Selected = mSelected
End Property


Property Get LastError() As String
    LastError = sError
End Property

Property Let Filename(v As String)
    sFilename = v
End Property

Property Get Filename() As String
    Filename = sFilename
End Property

Property Get FileTitle() As String
    FileTitle = sFileTitle
End Property

Property Let MaxSize(v As Long)
    mSize = v
End Property

Property Get MaxSize() As Long
    MaxSize = mSize
End Property

Property Get SelectedAnnotation() As Long
    SelectedAnnotation = lSelected
End Property

Public Function CreateThumbnailFromFile(Filename As String, MaxSize As Long) As Boolean
    sFilename = Filename
    smaxsize = MaxSize
    CreateThumbnailFromFile = CreateThumbnail
End Function

Public Function CreateThumbnailFromHandle(ByVal Handle As Long) As Boolean
Dim hTemp As Long

    If Handle <> 0 Then
        T = GetTickCount
        
        mOrigWidth = FreeImage_GetWidth(Handle)
        mOrigHeight = FreeImage_GetHeight(Handle)
        
        Dim temp As New CMemoryDC
        
        If mOrigHeight > 1024 Or mOrigWidth > 768 Then
            If mOrigHeight > mOrigWidth Then
                k = (4 * mSize) / mOrigHeight
            Else
                k = (4 * mSize) / mOrigWidth
            End If
            temp.Height = mOrigHeight * k
            temp.Width = mOrigWidth * k
        Else
            temp.Height = mOrigHeight
            temp.Width = mOrigWidth
        End If

        FreeImage_PaintDCEx temp.hDC, Handle, 0, 0, temp.Width, temp.Height, 0, 0, mOrigWidth, mOrigHeight
        
        hTemp = FreeImage_CreateFromDC(temp.hDC)
        
        hImage = FreeImage_MakeThumbnail(hTemp, mSize)
        Set temp = Nothing
        
        Debug.Print "Creating: " & Round((GetTickCount - T) / 1000, 2) & "s"
        
        If hImage <> 0 Then
            mWidth = FreeImage_GetWidth(hImage)
            mHeight = FreeImage_GetHeight(hImage)
            
            vZoom = mWidth / mOrigWidth
            
            CreateThumbnailFromHandle = True
            
            sFileTitle = Mid(sFilename, InStrRev(sFilename, "\") + 1)
            mHasThumbnail = True
        Else
            sError = "Error creating thumbnail"
        End If
    
        FreeImage_Unload hTemp
    Else
        sError = "Error opening handle"
    End If
End Function

Public Function CreateThumbnail() As Boolean
Dim hTemp As Long
Dim hTemp2 As Long

    If hImage <> 0 Then
        FreeImage_Unload hImage
        mHasThumbnail = False
        lCurPage = 1
    End If
    
    If Dir(sFilename) <> "" Then
        
        T = GetTickCount
        
        hTemp = FreeImage_Load(FreeImage_GetFIFFromFilename(sFilename), sFilename)
        
        Debug.Print "Loading: " & Round((GetTickCount - T) / 1000, 2) & "s"
        
        CreateThumbnail = CreateThumbnailFromHandle(hTemp)
        
        FreeImage_Unload hTemp
        
    Else
        sError = "File not found"
    End If

End Function

Private Sub DrawAnnotations()
Dim overlay As IOverlay
    For Each overlay In mCAnnotations
        overlay.Render Me
    Next
End Sub

Public Sub DrawThumbnail(hDC As Long, X As Long, Y As Long)
Dim mPen As LOGPEN
Dim lastpt As POINTAPI
    
    offsetX = X
    offsetY = Y
    
    mhDC = hDC
    
    If hImage <> 0 And mHasThumbnail Then
        FreeImage_PaintDCEx hDC, hImage, X, Y, mWidth, mHeight, 0, 0, mWidth, mHeight
        
        DrawAnnotations
    End If
End Sub

Private Function ScreenToImagePt(x1, y1)
Dim zfact As Single
    zfact = 100 / vZoom
        
    x1 = (x1 - offsetX) / vZoom
    y1 = (y1 - offsetY) / vZoom
    
    'UnRotatePoint x1, y1, mRotation
    
End Function

Public Sub MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)
Dim ctr As Long
' remove redundancies
    dx = X
    dy = Y

    px = X
    py = Y

    tx = X
    ty = Y

    If mParentControl.ShowNavigator Then
        With mNavigator
            If X - offsetX > .Left And X - offsetX < .Right And Y - offsetY > .Top And Y - offsetY < .Bottom Then
                mNavigatorSelected = True
            Else
                mNavigatorSelected = False
            End If
        End With
    Else
    
    Dim overlay As IOverlay
    
    For Each overlay In mCSelection
         overlay.Selected = False
    Next
    
    mCSelection.Clear
        
    If mCAnnotations.Count > 0 Then
       
       For Each overlay In mCAnnotations
            Dim hitType As HitTestEnum
            
            hitType = overlay.HitTest(X, Y, Me)
            
            If hitType <> HIT_NONE Then
                overlay.Selected = True
                'If hitType = HIT_CENTER Then
                    'RaiseEvent AnnotationClick(ctr)
                    'lSelected = ctr
                'End If
                mCSelection.Add overlay
                Exit For    ' avoid selecting below
            End If
            ctr = ctr + 1
       Next
       
       mCAnnotations.CallUpdate
    End If
    
    End If
    
    px = X
    py = Y
    
    sX = X
    sY = Y
End Sub

Public Sub MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)
Dim nBottom As Long
Dim nTop As Long
Dim nLeft As Long
Dim nRight As Long
Dim dx As Long
Dim dy As Long
Dim tRect As RECT
Dim overlay As IOverlay
    
    Select Case Button
    Case 1
        If mParentControl.ShowNavigator Then
            If mNavigatorSelected Then
                With mNavigator
                    dy = (Y - py)
                    dx = (X - px)
                    
                    nBottom = .Bottom + dy
                    nTop = .Top + dy
                    nLeft = .Left + dx
                    nRight = .Right + dx
                    
                    If nTop < 0 Or nBottom > mHeight Then dy = 0
                    If nLeft < 0 Or nRight > mWidth Then dx = 0
                    
                    .Top = .Top + dy
                    .Left = .Left + dx
                    .Right = .Right + dx
                    .Bottom = .Bottom + dy
                    
                    INavigator_Update
                    'If Not mParentControl Is Nothing Then
                    '    mParentControl.RaiseNavigatorChange .Left / vZoom, .Top / vZoom
                    'End If
                End With
            End If
        
        Else     'MOusemove for Annotation
        
            If mCSelection.Count > 0 Then
                For Each overlay In mCSelection
                    overlay.Move (X - px) / vZoom, (Y - py) / vZoom, Me
                Next

                mCSelection.CallUpdate
            
            Else
                mParentControl.DrawSelection = True
                mParentControl.SetSelection sX, sY, X, Y
            
            End If
        
        End If
        
        px = X
        py = Y
    
    Case 0
    
        If mCAnnotations.Count > 0 Then
            For Each overlay In mCAnnotations
                Dim hitType As HitTestEnum
                
                hitType = overlay.HitTest(X, Y, Me)
                
                Select Case hitType
                Case HitTestEnum.HIT_NONE, HitTestEnum.HIT_CENTER
                    mMousePointer = vbDefault
                    'Exit For    ' avoid selecting below
                
                Case HitTestEnum.HIT_TOP, HitTestEnum.HIT_BOTTOM
                    mMousePointer = vbSizeNS
                    Exit For    ' avoid selecting below
                
                Case HitTestEnum.HIT_LEFT, HitTestEnum.HIT_RIGHT
                    mMousePointer = vbSizeWE
                    Exit For    ' avoid selecting below
                
                Case HitTestEnum.HIT_LEFT + HitTestEnum.HIT_TOP, HitTestEnum.HIT_RIGHT + HitTestEnum.HIT_BOTTOM
                    mMousePointer = vbSizeNWSE
                    Exit For    ' avoid selecting below
                
                Case HitTestEnum.HIT_RIGHT + HitTestEnum.HIT_TOP, HitTestEnum.HIT_LEFT + HitTestEnum.HIT_BOTTOM
                    mMousePointer = vbSizeNESW
                    Exit For    ' avoid selecting below
                End Select
           Next
        Else
            mMousePointer = vbDefault
        End If
    
    End Select
End Sub

Public Sub MoveAnnotation(Index, X, Y)
'    With mAnnotations(Index).coords
'        If .Top + Y < 0 Or .Bottom + Y > mHeight Then Y = 0
'        If .Left + X < 0 Or .Right + X > mWidth Then X = 0
'
'        .Top = .Top + Y
'        .Left = .Left + X
'        .Right = .Right + X
'        .Bottom = .Bottom + Y
'
'        mParentControl.RaiseAnnotationChange Index, .Left / vZoom, .Top / vZoom, (.Right - .Left) / vZoom, (.Bottom - .Top) / vZoom
'    End With
End Sub

Public Sub MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)
    If Button = 1 Then
        If mCSelection.Count = 0 Then
            
            AddAnnotation (sX - offsetX) / vZoom, (sY - offsetY) / vZoom, (X - sX) / vZoom, (Y - sY) / vZoom, False
            
            mParentControl.DrawSelection = False
        End If
        
        ' check if changes were made....
        mCAnnotations.CallUpdate
    End If
End Sub

Private Sub Class_Initialize()
    mSize = 150
    Set mCAnnotations = New CAnnotations
    Set mCSelection = New CAnnotations
    Set mNavigatorEvents = New INavigatorEvents
    Set mNavigator = New CRect
End Sub

Private Sub Class_Terminate()
    Debug.Print "Destroying thumbnail... " & sFileTitle
    If hImage <> 0 Then
        FreeImage_Unload hImage
        mHasThumbnail = False
    End If
    
    Set mNavigator = Nothing

'    If Not mCSelection Is Nothing Then
'        mCSelection.Clear
'        Set mCSelection = Nothing
'    End If
'
'    If Not mCAnnotations Is Nothing Then
'        Dim overlay As Object
'        For Each overlay In mCAnnotations
'            Set overlay = Nothing
'        Next
'
'        mCAnnotations.Clear
'        Set mCAnnotations = Nothing
'    End If
        
    
End Sub

Property Get AnnotationCount() As Long
    AnnotationCount = mCAnnotations.Count
End Property

Public Sub GetAnnotation(ByVal Index As Long, Left As Long, Top As Long, Width As Long, Height As Long)
Dim thisAnnotation As CAnnotation
    
    Set thisAnnotation = mCAnnotations(Index + 1)
    
    With thisAnnotation
        Left = .Left / vZoom
        Top = .Top / vZoom
        Width = .Width / vZoom
        Height = .Height / vZoom
    End With
End Sub

Public Sub EditAnnotation(Index As Long, ByVal HitPoint As HitTestEnum, X As Long, Y As Long)
Dim thisAnnotation As IOverlay
    Set thisAnnotation = mCAnnotations(Index + 1)
    thisAnnotation.HitPoint = HitPoint
    thisAnnotation.Move X, Y, Me
End Sub

Public Sub RemoveAnnotation(ByVal Index As Long)
    mCAnnotations.RemoveAt Index + 1
    Set mCAnnotations(Index + 1) = Nothing
End Sub

Public Sub ClearAnnotations()
Dim thisAnnotation As CAnnotation
    
    mCSelection.Clear
    
    For Each thisAnnotation In mCAnnotations
        Set thisAnnotation = Nothing
    Next
    mCAnnotations.Clear
End Sub

Property Set Annotations(v As CAnnotations)
    Set mCAnnotations = v
End Property

Property Get Annotations() As CAnnotations
    Set Annotations = mCAnnotations
End Property

Private Property Let IContainer_DrawMode(RHS As Long)
    Debug.Print RHS
End Property

Private Property Let IContainer_ForeColor(RHS As Long)
    Debug.Print RHS
End Property

Private Property Get IContainer_hDC() As Long
    IContainer_hDC = mhDC
End Property

Private Property Get IContainer_Height() As Long
    IContainer_Height = mHeight / vZoom
End Property

Private Function ImageToScreen(v As CRect) As CRect
    Set ImageToScreen = New CRect

    ImageToScreen.Left = v.Left * vZoom + offsetX
    ImageToScreen.Top = v.Top * vZoom + offsetY
    ImageToScreen.Right = v.Right * vZoom + offsetX
    ImageToScreen.Bottom = v.Bottom * vZoom + offsetY
End Function

Private Function IContainer_ImageToScreen(v As CRect) As CRect
    Set IContainer_ImageToScreen = ImageToScreen(v)
End Function

Private Property Get IContainer_Page() As Long
    IContainer_Page = 1
End Property

Private Property Get IContainer_Rotation() As Long
    IContainer_Rotation = 0
End Property

Private Property Get IContainer_Width() As Long
    IContainer_Width = mWidth / vZoom
End Property

Private Property Let INavigator_Left(RHS As Long)
    mNavigator.Left = RHS
End Property

Private Property Get INavigator_Left() As Long
    INavigator_Left = mNavigator.Left
End Property

Private Sub INavigator_SetTarget(v As INavigatorEvents)
    Set mNavigatorEvents = v
    SetNavigator v.Left, v.Top, v.Width, v.Height
End Sub

Private Property Let INavigator_Top(RHS As Long)
    mNavigator.Top = RHS
End Property

Private Property Get INavigator_Top() As Long
    INavigator_Top = mNavigator.Top
End Property

Private Sub INavigator_Update()
    If Not mNavigatorEvents Is Nothing Then
        mNavigatorEvents.NavigatorUpdated mNavigator.Left / vZoom, mNavigator.Top / vZoom
        mNavigatorEvents.Redraw
    End If
End Sub

Private Sub INavigator_UpdateEx(Left As Long, Top As Long, Width As Long, Height As Long)
    mNavigator.Left = Left * vZoom
    mNavigator.Top = Top * vZoom
    mNavigator.Right = mNavigator.Left + Width * vZoom
    mNavigator.Bottom = mNavigator.Top + Height * vZoom
    mParentControl.Refresh
End Sub

Private Sub mCAnnotations_Updated()
    mParentControl.Refresh
End Sub

Private Sub mCSelection_Updated()
    mParentControl.Refresh
End Sub

Property Set Selection(v As CAnnotations)
    Set mCSelection = v
End Property

Property Get Selection() As CAnnotations
    Set Selection = mCSelection
End Property

