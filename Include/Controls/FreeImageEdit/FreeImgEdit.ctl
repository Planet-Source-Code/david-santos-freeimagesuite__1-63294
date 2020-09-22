VERSION 5.00
Begin VB.UserControl FreeImgEdit 
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4110
   KeyPreview      =   -1  'True
   ScaleHeight     =   3855
   ScaleWidth      =   4110
   ToolboxBitmap   =   "FreeImgEdit.ctx":0000
   Begin VB.PictureBox picDisplay 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   0
      ScaleHeight     =   3585
      ScaleWidth      =   3825
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
   Begin VB.HScrollBar HScrollBar 
      Enabled         =   0   'False
      Height          =   255
      LargeChange     =   50
      Left            =   0
      Max             =   0
      SmallChange     =   25
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3600
      Width           =   3855
   End
   Begin VB.VScrollBar VScrollBar 
      Enabled         =   0   'False
      Height          =   3615
      LargeChange     =   100
      Left            =   3840
      Max             =   0
      SmallChange     =   25
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "FreeImgEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements IContainer
Implements INavigatorEvents

Option Explicit

' TODO:
'    Faster rotation, and background updates when the image is large
'    Image being viewed before rotation must be equal to the image after rotation, the only difference is rotated

Private WithEvents HScroll1 As CLongScroll
Attribute HScroll1.VB_VarHelpID = -1
Private WithEvents VScroll1 As CLongScroll
Attribute VScroll1.VB_VarHelpID = -1
Dim pNavigator As INavigator
Private mAutoAnnotate As Boolean

'************************************************************************************
'*                          FreeImageEdit Internal Variables
'************************************************************************************

Dim mFileName As String
Dim mAutoRedraw As Boolean
Dim lCurPage As Long
Dim lLastPage As Long
Dim lPageCount As Long
Dim mRotation As Long
Dim lCurFilename As String
Dim zfact As Single

Dim hMultiBitmap As Long
Dim mBPP As Long

' handle to the in-memory image
Dim hImage As Long
Dim hCopy As Long

' stores the current width & height
Dim mImageWidth As Long
Dim mImageHeight As Long

' stores the (unrotated) width & height
Dim mAbsImageWidth As Long
Dim mAbsImageHeight As Long

Dim lResX As Long
Dim lResY As Long
Dim hBGBrush As Long

Dim lastHScroll As Long
Dim lastVScroll As Long

Dim dH As Long
Dim dV As Long

Dim mLineWidth As Long

Dim hSelectedBrush As Long
Dim hBlackBrush As Long
Dim bScrollEvents As Boolean

Dim dx As Long
Dim dy As Long

Dim lAnnotate As Long
Dim bMouseScroll As Boolean
Dim bUpdate As Boolean
Dim bEdit As Boolean
Dim bMove As Boolean
Dim lSelected As Long
Dim pX As Long
Dim pY As Long
Dim fmt As FREE_IMAGE_FORMAT
Dim bChanged As Boolean
Dim bFastSelection As Boolean
Dim bShowScrollBars As Boolean

Dim mSelectedColor As Long
Dim mAnnotationColor As Long
Dim mDefaultAnnotationFillColor As Long
Dim mDefaultAnnotationLineColor As Long
Dim bPreShowSelect As Boolean
Dim Weights(255) As Long

Dim mRubberBand As RubberBandMode

Dim offsetX As Long
Dim offsetY As Long

Dim mBuffer As CMemoryDC

Dim mScrollSizeX As Long
Dim mScrollSizeY As Long

'Dim mAnnotations() As tAnnotate
'Dim mAnnotationCount As Long

Dim mCursorX As Long
Dim mCursorY As Long
Dim bLocked As Boolean

Dim bShowOrientation As Boolean

Dim TPPX As Long
Dim TPPY As Long

Dim txDown As Long, tyDown As Long

Dim MFreeThumb As FreeThumb
Dim mAnnotationFilled As Boolean
Dim mDefaultAnnotationType As ANNOTATION_TYPE
Dim mShowNumber As Boolean

Dim mRotateDirection As RotateDirection

Dim mReflectRotationOnAnnotation As Boolean

Dim WithEvents mCAnnotations As CAnnotations
Attribute mCAnnotations.VB_VarHelpID = -1
Dim WithEvents mCSelection As CAnnotations
Attribute mCSelection.VB_VarHelpID = -1

Public Enum SplitDirection
    Split_Vertical = 0
    Split_Horizontal = 1
End Enum

Public Enum RotateDirection
    Rotate_Right = 1
    Rotate_Left = -1
    NoRotation = 0
End Enum

Public Enum RubberBandMode
    RUBBERBAND_BOX = 0
    RUBBERBAND_LINE = 1
End Enum

'************************************************************************************
'*                               FreeImageEdit Enums
'************************************************************************************
Public Enum AnnotationEditMode
    EDIT_NONE = 0
    EDIT_TOPLEFT = 1
    EDIT_TOPRIGHT = 2
    EDIT_BOTTOMLEFT = 3
    EDIT_BOTTOMRIGHT = 4
    EDIT_TOP = 5
    EDIT_RIGHT = 6
    EDIT_BOTTOM = 7
    EDIT_LEFT = 8
End Enum

Public Enum eFitTo
    FIT_ACTUAL_PIXELS = 0
    FIT_WIDTH = 1
    FIT_HEIGHT = 2
    FIT_BEST = 3
End Enum

'Private Enum BLIT_MODE
'    BLIT_DIRECT = 0
'    BLIT_BUFFER = 1
'End Enum


'Public Enum FREE_IMAGE_CONVERSION_FLAGS_VB
'   FICF_MONOCHROME = &H1
'   FICF_MONOCHROME_THRESHOLD = FICF_MONOCHROME
'   FICF_MONOCHROME_DITHER = &H3
'   FICF_GREYSCALE_4BPP = &H4
'   FICF_PALETTISED_8BPP = &H8
'   FICF_GREYSCALE_8BPP = FICF_PALETTISED_8BPP Or FICF_MONOCHROME
'   FICF_GREYSCALE = FICF_GREYSCALE_8BPP
'   FICF_RGB_15BPP = &HF
'   FICF_RGB_16BPP = &H10
'   FICF_RGB_24BPP = &H18
'   FICF_RGB_32BPP = &H20
'   FICF_RGB_ALPHA = FICF_RGB_32BPP
'   FICF_PREPARE_RESCALE = &H100
'   FICF_KEEP_UNORDERED_GREYSCALE_PALETTE = &H0
'   FICF_REORDER_GREYSCALE_PALETTE = &H1000
'End Enum
'#If False Then
'   Const FICF_MONOCHROME = &H1
'   Const FICF_MONOCHROME_THRESHOLD = FICF_MONOCHROME
'   Const FICF_MONOCHROME_DITHER = &H3
'   Const FICF_GREYSCALE_4BPP = &H4
'   Const FICF_PALETTISED_8BPP = &H8
'   Const FICF_GREYSCALE_8BPP = FICF_PALETTISED_8BPP Or FICF_MONOCHROME
'   Const FICF_GREYSCALE = FICF_GREYSCALE_8BPP
'   Const FICF_RGB_15BPP = &HF
'   Const FICF_RGB_16BPP = &H10
'   Const FICF_RGB_24BPP = &H18
'   Const FICF_RGB_32BPP = &H20
'   Const FICF_RGB_ALPHA = FICF_RGB_32BPP
'   Const FICF_PREPARE_RESCALE = &H100
'   Const FICF_KEEP_UNORDERED_GREYSCALE_PALETTE = &H0
'   Const FICF_REORDER_GREYSCALE_PALETTE = &H1000
'#End If


'************************************************************************************
'*                               FreeImageEdit Consts
'************************************************************************************


'************************************************************************************
'*                               FreeImageEdit Events
'************************************************************************************

Public Event AfterAnnotationCreate(ByVal Index As Long)
Public Event AnnotationCreate(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, ByVal Rotation As Long, ByVal Page As Long)
Public Event BeforeAnnotationCreate(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, _
                                        Cancel As Boolean, ByRef AnnotationType As ANNOTATION_TYPE, _
                                        ByRef FillColor As Long, ByRef LineColor As Long, ByRef Filled As Boolean)
                                        
Public Event AnnotationRemove(ByVal Index As Long)
Public Event AnnotationClick(ByVal Index As Long)
Public Event AnnotationChange(ByVal Index As Long, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, Cancel As Boolean)

Public Event ZoomChange(ByVal ZoomFactor As Single)
Public Event PageChange(ByVal LastPage As Long, ByVal NewPage As Long)
Public Event ScrollChange(ByVal X As Long, ByVal Y As Long)

Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'#If DEPRECATED_EVENTS_INCLUDED Then
'Public Event PreDraw(ByVal hDC As Long)
'Public Event PostDraw(ByVal hDC As Long)
'Public Event PostDrawAnnotation(ByVal hDC As Long, ByVal Index As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, ByVal Rotation As Long)
'#End If

Public Event AfterDisplay(ByVal hdc As Long)
Attribute AfterDisplay.VB_Description = "Raised after all annotations are drawn. Can be used to draw user information anywhere on the entire window."
Public Event AfterAnnotationDraw(ByVal hdc As Long, ByVal Index As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, ByVal Rotation As Long)
Attribute AfterAnnotationDraw.VB_Description = "Raised after each annotation is drawn. Can be used to draw user information over selective annotations."

Public Property Get HScrollValue() As Long
    HScrollValue = HScroll1.Value
End Property

Public Property Get VScrollValue() As Long
    VScrollValue = VScroll1.Value
End Property

Public Property Let ReflectRotationOnAnnotation(v As Boolean)
    mReflectRotationOnAnnotation = v
End Property

Public Property Get ReflectRotationOnAnnotation() As Boolean
    ReflectRotationOnAnnotation = mReflectRotationOnAnnotation
    PropertyChanged "ReflectRotationOnAnnotation"
End Property

Property Let ScrollUpdates(v As Boolean)
    bUpdate = v
End Property

Property Get ScrollUpdates() As Boolean
    ScrollUpdates = bUpdate
End Property

Property Get VisibleWidth() As Long
    If offsetX = 0 Then
        VisibleWidth = picDisplay.Width / TPPY * zfact
    Else
        VisibleWidth = mImageWidth
    End If
End Property

Property Get VisibleHeight() As Long
    If offsetY = 0 Then
        VisibleHeight = picDisplay.Height / TPPX * zfact
    Else
        VisibleHeight = mImageHeight
    End If
End Property

Property Let ScrollEvents(v As Boolean)
    bScrollEvents = v
End Property

Property Get ScrollEvents() As Boolean
    ScrollEvents = bScrollEvents
End Property

Property Get LineWidth() As Long
    LineWidth = mLineWidth
End Property

Property Let LineWidth(v As Long)
    mLineWidth = v
End Property

Property Get AnnotationCount() As Long
Attribute AnnotationCount.VB_MemberFlags = "400"
    AnnotationCount = mCAnnotations.Count
End Property

'Read only???
'Property Let AnnotationType(Index, Mode As ANNOTATION_TYPE)
'    'mAnnotations(index).Type = Mode
'End Property

Property Get AnnotationType(Index) As ANNOTATION_TYPE
    Dim thisAnnotation As IOverlay
    Set thisAnnotation = mCAnnotations(Index + 1)
    AnnotationType = thisAnnotation.AnnotationType
End Property

Property Let FillColor(Index, v As Long)
    Dim thisOverlay As IOverlay
    Set thisOverlay = mCAnnotations(Index + 1)
    If thisOverlay.AnnotationType = fieFilledRect Or thisOverlay.AnnotationType = fieHollowRect Then
        Dim thisAnnotation As CAnnotation
        Set thisAnnotation = mCAnnotations(Index + 1)
        thisAnnotation.FillColor = v
    End If
End Property

Property Get FillColor(Index) As Long
    Dim thisOverlay As IOverlay
    Set thisOverlay = mCAnnotations(Index + 1)
    If thisOverlay.AnnotationType = fieFilledRect Or thisOverlay.AnnotationType = fieHollowRect Then
        Dim thisAnnotation As CAnnotation
        Set thisAnnotation = mCAnnotations(Index + 1)
        FillColor = thisAnnotation.FillColor
    End If
End Property

Property Let AnnotationLeft(Index, v As Long)
Attribute AnnotationLeft.VB_MemberFlags = "400"
    Dim thisAnnotation As IOverlay
    Set thisAnnotation = mCAnnotations(Index + 1)
    thisAnnotation.Left = v
End Property

Property Get AnnotationLeft(Index) As Long
    Dim thisAnnotation As IOverlay
    Set thisAnnotation = mCAnnotations(Index + 1)
    AnnotationLeft = thisAnnotation.Left
End Property

Property Let AnnotationTop(Index, v As Long)
Attribute AnnotationTop.VB_MemberFlags = "400"
    Dim thisAnnotation As IOverlay
    Set thisAnnotation = mCAnnotations(Index + 1)
    thisAnnotation.Top = v
End Property

Property Get AnnotationTop(Index) As Long
    Dim thisAnnotation As IOverlay
    Set thisAnnotation = mCAnnotations(Index + 1)
    AnnotationTop = thisAnnotation.Top
End Property

Property Get AnnotationWidth(Index) As Long
Attribute AnnotationWidth.VB_MemberFlags = "400"
    Dim thisAnnotation As IOverlay
    Set thisAnnotation = mCAnnotations(Index + 1)
    AnnotationWidth = thisAnnotation.Width
End Property

Property Let AnnotationWidth(Index, v As Long)
    Dim thisAnnotation As IOverlay
    Set thisAnnotation = mCAnnotations(Index + 1)
    thisAnnotation.Width = v
End Property

Property Get AnnotationHeight(Index) As Long
Attribute AnnotationHeight.VB_MemberFlags = "400"
    Dim thisAnnotation As IOverlay
    Set thisAnnotation = mCAnnotations(Index + 1)
    AnnotationHeight = thisAnnotation.Height
End Property

Property Let AnnotationHeight(Index, v As Long)
    Dim thisAnnotation As IOverlay
    Set thisAnnotation = mCAnnotations(Index + 1)
    thisAnnotation.Height = v
End Property

Property Get AnnotationRotation(Index) As Long
Attribute AnnotationRotation.VB_MemberFlags = "400"
    Dim thisAnnotation As IOverlay
    Set thisAnnotation = mCAnnotations(Index + 1)
    AnnotationRotation = thisAnnotation.Rotation
End Property

Property Let AnnotationRotation(Index, v As Long)
    Dim thisAnnotation As IOverlay
    Set thisAnnotation = mCAnnotations(Index + 1)
    thisAnnotation.Rotation = v
End Property

Property Get AnnotationLineColor(Index) As ColorConstants
Attribute AnnotationLineColor.VB_MemberFlags = "400"
'    Dim thisOverlay As IOverlay
'    Set thisOverlay = mCAnnotations(Index + 1)
'    If thisAnnotation.AnnotationType = fieFilledRect Or thisAnnotation.AnnotationType = fieHollowRect Then
'        Dim thisAnnotation As CAnnotation
'        Set thisAnnotation = mCAnnotations(Index + 1)
'        FillColor = thisAnnotation.FillColor
'    End If
    Err.Raise vbObjectError, "AnnotationLineColor", "Not yet implemented"
End Property

Property Let AnnotationLineColor(Index, color As Long)
'    Dim thisOverlay As IOverlay
'    Set thisOverlay = mCAnnotations(Index + 1)
'    If thisAnnotation.AnnotationType = fieFilledRect Or thisAnnotation.AnnotationType = fieHollowRect Then
'        Dim thisAnnotation As CLine
'        Set thisAnnotation = mCAnnotations(Index + 1)
'        FillColor = thisAnnotation.LineWidth
'    End If
    Err.Raise vbObjectError, "AnnotationLineColor", "Not yet implemented"
End Property

Property Let AutoRedraw(Redraw As Boolean)
    mAutoRedraw = Redraw
End Property

Property Get AutoRedraw() As Boolean
    AutoRedraw = mAutoRedraw
End Property

Property Get DefaultAnnotationType() As ANNOTATION_TYPE
    DefaultAnnotationType = mDefaultAnnotationType
End Property

Property Let DefaultAnnotationType(ByVal v As ANNOTATION_TYPE)
    mDefaultAnnotationType = v
    PropertyChanged "DefaultAnnotationType"
End Property

Property Get DefaultAnnotationFillColor() As OLE_COLOR
    DefaultAnnotationFillColor = mDefaultAnnotationFillColor
    PropertyChanged "DefaultAnnotationFillColor"
End Property

Property Let DefaultAnnotationFillColor(ByVal vDefaultAnnotationFillColor As OLE_COLOR)
    mDefaultAnnotationFillColor = vDefaultAnnotationFillColor
End Property

Property Get DefaultAnnotationLineColor() As OLE_COLOR
    DefaultAnnotationLineColor = mDefaultAnnotationLineColor
    PropertyChanged "DefaultAnnotationLineColor"
End Property

Property Let DefaultAnnotationLineColor(ByVal vDefaultAnnotationLineColor As OLE_COLOR)
    mDefaultAnnotationLineColor = vDefaultAnnotationLineColor
End Property

Property Get AnnotationFilled() As Boolean
    AnnotationFilled = mAnnotationFilled
    PropertyChanged "AnnotationFilled"
End Property

Property Let AnnotationFilled(ByVal vAnnotationFilled As Boolean)
    mAnnotationFilled = vAnnotationFilled
End Property

Property Get ShowNumber() As Boolean
    ShowNumber = mShowNumber
End Property

Property Let ShowNumber(ByVal vShowNumber As Boolean)
    mShowNumber = vShowNumber
End Property

Property Get ImageHandle() As Long
    ImageHandle = hImage
End Property

Property Let Image(Filename)
Attribute Image.VB_MemberFlags = "400"
    On Error GoTo Image_Error
    Dim tempImage As Long
    Dim newimage As Long

'    If lCurFilename = Filename Then
'        Exit Property
'    Else
'        lCurFilename = Filename
'    End If

    mRotateDirection = NoRotation

    If hMultiBitmap <> 0 Then
        FreeImage_CloseMultiBitmap hMultiBitmap
        hMultiBitmap = 0
    End If

    If hImage <> 0 Then
        FreeImage_Unload hImage
        hImage = 0
    End If

    If hCopy <> 0 Then
        FreeImage_Unload hCopy
        hCopy = 0
    End If

    ' blank assumes we want to unload the image from memory
    VScrollBar.Enabled = False
    HScrollBar.Enabled = False
    If Filename = "" Then Exit Property

    mFileName = Filename

    ' Set format based on extension
    fmt = FreeImage_GetFIFFromFilename(mFileName)

    ' check for multi-page formats
    If (fmt = FIF_TIFF) Or (fmt = FIF_GIF) Then

        ' try to load the file
        hMultiBitmap = FreeImage_OpenMultiBitmap(fmt, Filename, CLng(False), CLng(False))

        If hMultiBitmap <> 0 Then
            lPageCount = FreeImage_GetPageCount(hMultiBitmap)
            ' get the first page
            lCurPage = 0
            tempImage = FreeImage_LockPage(hMultiBitmap, lCurPage)
            ' make a copy of this page so we can rotate, etc...
            newimage = FreeImage_Clone(tempImage)
            ' return the page to the bitmap
            FreeImage_UnlockPage hMultiBitmap, tempImage, CLng(False)
        End If

    Else
        ' single page image
        lPageCount = 1
        newimage = FreeImage_Load(fmt, mFileName)

    End If

    '_Load or _OpenMultiBitmap will set fmt to FIF_UNKNOWN
    'Select Case fmt
    'Case FIF_UNKNOWN
    '    Err.Raise vbObjectError + 6, "FreeImageEdit.Image (Property)", "FreeImageEdit.Image (Property): Sorry, I don't recognize this file."
    'Case Else

    If newimage <> 0 Then

        lResX = FreeImage_GetDotsPerMeterX(newimage)
        lResY = FreeImage_GetDotsPerMeterY(newimage)
        
        UpdateImgBuffer newimage

        mAbsImageWidth = mImageWidth
        mAbsImageHeight = mImageHeight

        mScrollSizeX = mImageWidth * 0.05
        mScrollSizeY = mImageHeight * 0.05

        UserControl_Resize
    
    Else
        Err.Raise vbObjectError + 6, "FreeImageEdit.Image (Property)", "Error opening image"

    End If

    'End Select

    ' reset rotation
    mRotation = 0

    On Error GoTo 0
    Exit Property

Image_Error:
    Dim omf As Long
    Err.Raise Err.Number, "Property Image Let in FreeImgEdit", Err.Description

End Property

Property Get Image()
    Image = mFileName
End Property

Property Get ImageWidth() As Long
Attribute ImageWidth.VB_MemberFlags = "400"
    ImageWidth = mImageWidth
End Property

Property Get ImageHeight() As Long
Attribute ImageHeight.VB_MemberFlags = "400"
    ImageHeight = mImageHeight
End Property

Property Get MouseScroll() As Boolean
Attribute MouseScroll.VB_MemberFlags = "400"
    MouseScroll = bMouseScroll
End Property

' Sets the mouse scrolling feature
' enabling this will disable highlight creation
' as mouse click-and-drag will be used for scrolling
Property Let MouseScroll(Mode As Boolean)
    bMouseScroll = Mode
End Property

Property Get Page() As Long
Attribute Page.VB_MemberFlags = "400"
    Page = lCurPage + 1
End Property

Property Let Page(lPage As Long)
Dim newimage As Long
Dim tempImage As Long
Dim lCount As Long
Dim oldPage As Long
Dim mCurrPage As Long
    
    mRotateDirection = NoRotation
    
    mCurrPage = lCurPage
    oldPage = lCurPage
    lCurPage = lPage - 1

    If lCurPage > lPageCount - 1 Then lCurPage = lPageCount - 1
    If lCurPage < 0 Then lCurPage = 0

    'If mCurrPage <> lCurPage Then
        If hMultiBitmap <> 0 Then
            tempImage = FreeImage_LockPage(hMultiBitmap, lCurPage)
            newimage = FreeImage_Clone(tempImage)
            FreeImage_UnlockPage hMultiBitmap, tempImage, CLng(False)
            UpdateImgBuffer newimage
            mAbsImageWidth = mImageWidth
            mAbsImageHeight = mImageHeight
            mRotation = 0
        End If
        'lLastPage = lCurPage
    'End If
    RaiseEvent PageChange(oldPage + 1, lPage)
End Property

Property Get PageCount() As Long
Attribute PageCount.VB_MemberFlags = "400"
    PageCount = lPageCount
End Property

Property Get Rotation() As Long
Attribute Rotation.VB_MemberFlags = "400"
    Rotation = mRotation
End Property

Property Let ScrollX(pos As Long)
Attribute ScrollX.VB_MemberFlags = "400"
    If pos > HScroll1.Max Then pos = HScroll1.Max
    If pos < 0 Then pos = 0
    'bUpdate = False
    HScroll1.Value = pos
    'bUpdate = True
End Property

Property Let ScrollY(pos As Long)
Attribute ScrollY.VB_MemberFlags = "400"
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

Property Let ScrollSizeX(lSize As Long)
Attribute ScrollSizeX.VB_MemberFlags = "400"
    mScrollSizeX = lSize
End Property

Property Let ScrollSizeY(lSize As Long)
Attribute ScrollSizeY.VB_MemberFlags = "400"
    mScrollSizeY = lSize
End Property

Property Get ScrollSizeX() As Long
    ScrollSizeX = mScrollSizeX
End Property

Property Get ScrollSizeY() As Long
    ScrollSizeY = mScrollSizeY
End Property

'************************************************************************************
'*                                 FreeImageEdit Properties
'************************************************************************************
Property Get SelectedColor() As OLE_COLOR
Attribute SelectedColor.VB_Description = "Sets the border color used when displaying a selected annotation"
    SelectedColor = mSelectedColor
End Property

Property Let SelectedColor(ByVal color As OLE_COLOR)
    mSelectedColor = color
End Property

Property Get BackColor() As OLE_COLOR
    BackColor = picDisplay.BackColor
End Property

Property Let BackColor(ByVal color As OLE_COLOR)
    picDisplay.BackColor = color
End Property

Property Get SelectedAnnotation() As Long
Attribute SelectedAnnotation.VB_MemberFlags = "400"
    SelectedAnnotation = lSelected
End Property

Property Let SelectedAnnotation(Index As Long)
Dim i As Long
Dim thisAnnotation As IOverlay

    lSelected = Index
    'If lSelected < 0 Then lSelected = 0
    'If lSelected > mCAnnotations.Count - 1 Then lSelected = mCAnnotations.Count - 1

    mCSelection.Updates = False
    
    For Each thisAnnotation In mCSelection
        thisAnnotation.Selected = False
    Next
    
    mCSelection.Clear
    
    Set thisAnnotation = mCAnnotations(Index + 1)
    thisAnnotation.Selected = True

    mCSelection.Add thisAnnotation

    mCSelection.Updates = True
    mCSelection.CallUpdate

End Property

Property Let Zoom(fZoom As Single)
Attribute Zoom.VB_MemberFlags = "400"
    If fZoom > 2400 Then fZoom = 2400
    If fZoom < 2 Then fZoom = 2

    zfact = 100 / CSng(fZoom)
    
    RaiseEvent ZoomChange(zfact)
    
    UserControl_Resize
End Property

Property Get Zoom() As Single
    Zoom = 100 / zfact
End Property

Property Get hdc() As Long
Attribute hdc.VB_MemberFlags = "400"
    hdc = picDisplay.hdc
End Property

Property Get ShowOrientation() As Boolean
    ShowOrientation = bShowOrientation
End Property

Property Let ShowOrientation(Mode As Boolean)
    bShowOrientation = Mode
End Property

Property Let ScrollBarsVisible(Mode As Boolean)
    bShowScrollBars = Mode
    UserControl_Resize
    PropertyChanged "ScrollBarsVisible"
End Property

Property Get ScrollBarsVisible() As Boolean
    ScrollBarsVisible = bShowScrollBars
End Property

Property Get AutoAnnotate() As Boolean
    AutoAnnotate = mAutoAnnotate
End Property

Property Let AutoAnnotate(v As Boolean)
    mAutoAnnotate = v
End Property

Public Sub SplitAnnotation(Index As Long, Direction As SplitDirection)
Dim thisAnnotation As CAnnotation
Dim newAnnotation As CAnnotation
Dim newRect As New CRect
    
    Set thisAnnotation = mCAnnotations(Index + 1)
    
    If Direction = Split_Vertical Then
        thisAnnotation.Width = thisAnnotation.Width / 2
        
        Set newAnnotation = thisAnnotation.GetCopy
        
        newAnnotation.Left = thisAnnotation.Left + thisAnnotation.Width
        newAnnotation.Width = thisAnnotation.Width
    Else
        thisAnnotation.Height = thisAnnotation.Height / 2
        
        Set newAnnotation = thisAnnotation.GetCopy
        
        newAnnotation.Top = thisAnnotation.Top + thisAnnotation.Height
        newAnnotation.Height = thisAnnotation.Height
    End If
    
    mCAnnotations.Add newAnnotation
    
End Sub

Property Get CursorX() As Long
Attribute CursorX.VB_MemberFlags = "400"
    CursorX = mCursorX
End Property

Property Get CursorY() As Long
Attribute CursorY.VB_MemberFlags = "400"
    CursorY = mCursorY
End Property

Property Let Locked(Mode As Boolean)
Attribute Locked.VB_Description = "Sets whether the user can create/edit/remove annotations interactively during run-time"
    bLocked = Mode
End Property

Property Get Locked() As Boolean
    Locked = bLocked
    PropertyChanged "Locked"
End Property

Property Set BackgroundImage(ByVal BGPic As StdPicture)
    Set picDisplay.Picture = BGPic
    PropertyChanged "BackgroundImage"
End Property

Property Get BackgroundImage() As StdPicture
    Set BackgroundImage = picDisplay.Picture
End Property

Public Sub SetNavigator(Navigator As INavigator)
    Set pNavigator = Navigator
    Navigator.SetTarget Me
End Sub

Private Property Get INavigatorEvents_Height() As Long
    INavigatorEvents_Height = VisibleHeight
End Property

Private Property Get INavigatorEvents_Left() As Long
    INavigatorEvents_Left = ScrollX
End Property

Private Sub INavigatorEvents_NavigatorUpdated(Left As Long, Top As Long)
    ScrollEvents = False
    ScrollUpdates = False
    ScrollX = Left / zfact
    ScrollY = Top / zfact
    ScrollUpdates = True
    ScrollEvents = True
End Sub

Private Sub INavigatorEvents_Redraw()
    Debug.Print "INavigatorEvents_Redraw"
    Display
End Sub

Private Property Get INavigatorEvents_Top() As Long
    INavigatorEvents_Top = ScrollY
End Property

Private Property Get INavigatorEvents_Width() As Long
    INavigatorEvents_Width = VisibleWidth
End Property

Private Property Get INavigatorEvents_Zoom() As Single
    INavigatorEvents_Zoom = zfact
End Property

Private Sub UserControl_Initialize()
Dim i As Long, n As Long
Dim nbits As Byte

    TPPX = Screen.TwipsPerPixelX
    TPPY = Screen.TwipsPerPixelY
    
    'mLineWidth = 2
    bScrollEvents = True
    
    zfact = 1
    Set mBuffer = New CMemoryDC
    bUpdate = True
    lSelected = -1
    bShowOrientation = True

    hBlackBrush = CreateSolidBrush(RGB(255, 0, 0))
    hSelectedBrush = CreateSolidBrush(RGB(0, 0, 255))

    mScrollSizeX = 10
    mScrollSizeY = 10


    For i = 0 To 255
        nbits = 0
        For n = 0 To 7
            nbits = nbits + (i And (2 ^ n)) \ (2 ^ n)
        Next
        Weights(i) = nbits
    Next

    'mSelectedColor = RGB(0, 0, 255)
    mDefaultAnnotationLineColor = RGB(255, 0, 0)

    
    hBGBrush = CreateSolidBrush(picDisplay.BackColor)
    
    'mBlitMode = BLIT_DIRECT
    mRotation = 0

    'FreeImage_InitErrorHandler
    Set HScroll1 = New CLongScroll
    Set HScroll1.Client = HScrollBar
    
    Set VScroll1 = New CLongScroll
    Set VScroll1.Client = VScrollBar

    Set mCAnnotations = New CAnnotations
    Set mCSelection = New CAnnotations
    
    mAutoAnnotate = True
    mRubberBand = RUBBERBAND_BOX
End Sub

Private Sub UserControl_InitProperties()
    mSelectedColor = RGB(0, 0, 255)
    bShowScrollBars = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    bLocked = PropBag.ReadProperty("Locked", False)
    bShowScrollBars = PropBag.ReadProperty("ScrollBarsVisible", True)
    mSelectedColor = PropBag.ReadProperty("SelectedColor", RGB(0, 0, 255))
    picDisplay.BackColor = PropBag.ReadProperty("BackColor", RGB(128, 128, 128))
    Set picDisplay.Picture = PropBag.ReadProperty("BackgroundImage", Nothing)
    mLineWidth = PropBag.ReadProperty("LineWidth", 2)
    mAnnotationFilled = PropBag.ReadProperty("AnnotationFilled", False)
    mShowNumber = PropBag.ReadProperty("ShowNumber", False)
    mDefaultAnnotationType = PropBag.ReadProperty("DefaultAnnotationType", fieHollowRect)
    mRubberBand = PropBag.ReadProperty("Rubberband", RUBBERBAND_BOX)
    
    mDefaultAnnotationFillColor = PropBag.ReadProperty("DefaultAnnotationFillColor", RGB(255, 0, 255))
    
    mDefaultAnnotationLineColor = PropBag.ReadProperty("DefaultAnnotationLineColor", RGB(255, 0, 0))
    mReflectRotationOnAnnotation = PropBag.ReadProperty("ReflectRotationOnAnnotation", True)
    
'    If Not MFreeThumb Is Nothing Then
'        MFreeThumb = PropBag.ReadProperty("FreeThumbEdit", Nothing)
'    End If
    UserControl_Resize
End Sub

Private Sub UpdateWindowEvent()
    If Not pNavigator Is Nothing Then
        pNavigator.UpdateEx ScrollX * zfact, ScrollY * zfact, VisibleWidth, VisibleHeight
    End If
End Sub

Private Sub UserControl_Resize()
    Dim hMax As Long
    Dim vMax As Long
    Dim RemoveList() As Long
    Dim RemoveListCount As Long
    Dim picWidth As Long
    Dim picHeight As Long
    Dim mClientWidth As Long
    Dim mClientHeight As Long
    Dim tRect As RECT
    Dim vfact As Single
    Dim hfact As Single
    Dim mprev_HMax As Long
    Dim mprev_VMax As Long
    Dim mVScrollValue As Long
    Dim mHScrollValue As Long
    Dim mPercentVertical As Double
    Dim mPercentHorizontal As Double
    
    
    If UserControl.Height < 300 Then Exit Sub
    If UserControl.Width < 300 Then Exit Sub

    If bShowScrollBars Then
        HScrollBar.Height = 255
        VScrollBar.Width = 255
        HScrollBar.Visible = True
        VScrollBar.Visible = True
    Else
        HScrollBar.Visible = False
        VScrollBar.Visible = False
        HScrollBar.Height = 0
        VScrollBar.Width = 0
    End If

    picDisplay.Top = 0
    picDisplay.Left = 0
    HScrollBar.Left = 0
    VScrollBar.Top = 0
    
    mClientWidth = UserControl.Width - VScrollBar.Width
    mClientHeight = UserControl.Height - HScrollBar.Height
    
    picDisplay.Width = mClientWidth
    HScrollBar.Width = mClientWidth
    VScrollBar.Left = mClientWidth
    
    picDisplay.Height = mClientHeight
    VScrollBar.Height = mClientHeight
    HScrollBar.Top = mClientHeight

    dH = 0
    dV = 0

    mprev_HMax = HScroll1.Max
    mprev_VMax = VScroll1.Max
    mVScrollValue = VScroll1.Value
    mHScrollValue = HScroll1.Value
    

    If hImage <> 0 Then
        
        vMax = (mImageHeight / zfact) - (mClientHeight / TPPY)
        If vMax < 0 Then vMax = 0

        hMax = (mImageWidth / zfact) - (mClientWidth / TPPX)
        If hMax < 0 Then hMax = 0

        ' center image if zoomed out
        If hMax = 0 Then offsetX = (mClientWidth / TPPX - mImageWidth / zfact) / 2 Else offsetX = 0
        If vMax = 0 Then offsetY = (mClientHeight / TPPY - mImageHeight / zfact) / 2 Else offsetY = 0

        ' disable scroll updates
        
        bUpdate = False

        HScrollBar.Enabled = Not (hMax = 0)
        VScrollBar.Enabled = Not (vMax = 0)

        If VScroll1.Max > 0 Then vfact = VScroll1.Value / VScroll1.Max
        If HScroll1.Max > 0 Then hfact = HScroll1.Value / HScroll1.Max

        ScrollEvents = False
        VScroll1.Max = vMax
        HScroll1.Max = hMax
        
        
        
        If VScroll1.Value <> vfact * vMax Then VScroll1.Value = vfact * vMax
        If HScroll1.Value <> hfact * hMax Then HScroll1.Value = hfact * hMax
        
        'VScroll1.Value = vfact * vScroll1.Max
        'HScroll1.Value = hfact * HScroll1.Max
        
        
        'instead of using the ScrollEvents property, we set the variable the holds the property value
        'this optimizes speed
        bScrollEvents = True
        'ScrollEvents = True
        
        ' enable scroll updates
        bUpdate = True
        
        UpdateWindowEvent
        
        RaiseEvent ScrollChange(HScroll1.Value * zfact, VScroll1.Value * zfact)
        
        picWidth = mClientWidth / TPPX
        picHeight = mClientHeight / TPPY
    
        mBuffer.Height = picHeight
        mBuffer.Width = picWidth
        
        tRect.Top = 0
        tRect.Left = 0
        tRect.Right = picWidth
        tRect.Bottom = picHeight
        
        FillRect mBuffer.hdc, tRect, hBGBrush
        
        ' update image buffer with scaled image
        If offsetY > 0 Then
            'Debug.Print "Rotation: " & mRotation
            FreeImage_PaintDCEx mBuffer.hdc, hImage, offsetX, -picHeight + offsetY + mImageHeight / zfact, picWidth, picHeight, HScroll1.Value * zfact, (VScroll1.Max - VScroll1.Value) * zfact, picWidth * zfact, picHeight * zfact, DM_DRAW_DEFAULT, ROP_SRCCOPY
        
        Else
            'Debug.Print "Rotation: " & mRotation & "    mImageWidth = " & mImageWidth & "    mImageHeight = " & mImageHeight & "    mClientWidth = " & mClientWidth & "    mClientHeight = " & mClientHeight & "    HorizontalScroll = " & HScroll1.Value & " VerticalScroll = " & VScroll1.Value
            FreeImage_PaintDCEx mBuffer.hdc, hImage, offsetX, 0, picWidth, picHeight, HScroll1.Value * zfact, (VScroll1.Max - VScroll1.Value) * zfact, picWidth * zfact, picHeight * zfact, DM_DRAW_DEFAULT, ROP_SRCCOPY
        
        End If
                    
        If mHScrollValue = 0 And mprev_HMax = 0 Then
            mPercentVertical = 0
        Else
            mPercentVertical = mHScrollValue / mprev_HMax
        End If
        
        If mVScrollValue = 0 And mprev_VMax = 0 Then
            mPercentHorizontal = 0
        Else
            mPercentHorizontal = mVScrollValue / mprev_VMax
        End If
                    
        Select Case mRotateDirection
        Case Rotate_Right
            VScroll1.Value = mPercentVertical * vMax
            HScroll1.Value = hMax - (mPercentHorizontal * hMax)
            mRotateDirection = NoRotation
        Case Rotate_Left
            VScroll1.Value = vMax - (mPercentVertical * vMax)
            HScroll1.Value = (mPercentHorizontal * hMax)
            mRotateDirection = NoRotation
        End Select
    
    Else
        
        ' reset change in scroll
        picWidth = mClientWidth / TPPX
        picHeight = mClientHeight / TPPY
    
        mBuffer.Height = picHeight
        mBuffer.Width = picWidth
        
        tRect.Top = 0
        tRect.Left = 0
        tRect.Right = picWidth
        tRect.Bottom = picHeight
        
        FillRect mBuffer.hdc, tRect, hBGBrush
    End If

End Sub

Private Sub UserControl_Terminate()
    Debug.Print "Destroying viewer"
    If hImage <> 0 Then
        FreeImage_Unload hImage
        hImage = 0
    End If

    If hMultiBitmap <> 0 Then
        FreeImage_CloseMultiBitmap hMultiBitmap
        hMultiBitmap = 0
    End If

    If hCopy <> 0 Then
        FreeImage_Unload hCopy
        hCopy = 0
    End If

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
'
    DeleteObject hBlackBrush
    DeleteObject hBGBrush
    DeleteObject hSelectedBrush
    
    Set mBuffer = Nothing
End Sub

Public Function Save() As Boolean
    Save = FreeImage_Save(fmt, hImage, mFileName) = 0
End Function

Public Function SaveAs(Filename As String, Optional Options As Long) As Boolean
    If InStr(Filename, ".") = 0 Then
        ' do something?
        Err.Raise vbObjectError + 101, "FreeImgEdit.SaveAs", "No extension specified"
    
    Else
        Select Case UCase(Mid(Filename, InStrRev(Filename, ".") + 1))
        Case "TIF", "TIFF"
            fmt = FIF_TIFF
        Case "BMP"
            fmt = FIF_BMP
        Case "GIF"
            fmt = FIF_GIF
        Case "JPG", "JPEG"
            fmt = FIF_JPEG
        Case "TGA"
            fmt = FIF_TARGA
        Case "PNG"
            fmt = FIF_PNG
        End Select
    
    End If

    FreeImage_SetDotsPerMeterX hImage, lResX
    FreeImage_SetDotsPerMeterY hImage, lResY

    SaveAs = FreeImage_Save(fmt, hImage, Filename, Options) = 0
    If SaveAs Then mFileName = Filename
End Function

' clear background
' faster than picDisplay.Cls
Public Sub ClearDisplay()
    Dim hBrush As Long
    Dim oldBrush As Long

    hBrush = CreateSolidBrush(RGB(96, 96, 96))
    oldBrush = SelectObject(picDisplay.hdc, hBrush)
    Rectangle picDisplay.hdc, 0, 0, picDisplay.Width / TPPX, picDisplay.Height / TPPY
    SelectObject picDisplay.hdc, oldBrush
    DeleteObject hBrush

End Sub

Private Sub DrawSelection()
    'zfact = 100 / mZoom

'    If lSelected > -1 And mAnnotations(lSelected).Page = lCurPage Then
'        tRect = mAnnotations(lSelected).coords
'
'        'tRect = RotateRect(tRect, mRotation)
'        tRect = ImageToScreen(tRect)
'
'        lCurBrush = hSelectedBrush
'        hPen = CreatePen(PS_SOLID, 2&, mSelectedColor)
'        picDisplay.ForeColor = RGB(0, 0, 255)
'
'        picDisplay.DrawMode = vbInvert
'
'        hOldObj = SelectObject(picDisplay.hDC, hPen)
'
'        MoveToEx picDisplay.hDC, tRect.Left, tRect.Top, oldPoint
'        LineTo picDisplay.hDC, tRect.Right, tRect.Top
'        LineTo picDisplay.hDC, tRect.Right, tRect.Bottom
'        LineTo picDisplay.hDC, tRect.Left, tRect.Bottom
'        LineTo picDisplay.hDC, tRect.Left, tRect.Top
'
'        picDisplay.Refresh
'
'        ' Get rid of pen
'        SelectObject picDisplay.hDC, hOldObj
'        DeleteObject hPen
'    End If
End Sub

Private Sub NumberAnnotation(overlay As IOverlay, Index As Long)
Dim mTOut As Long
Dim sTemp As String
Dim tRect As RECT
Dim xc As Long
Dim yc As Long

mShowNumber = True
    If mShowNumber = True Then
        ' draw annotation number
    
        sTemp = Str(Index + 1)
    
        On Error Resume Next
        Dim tWidth As Long
        Dim tTop  As Long
        Dim tLeft As Long
        Dim tHeight As Long
        Dim mHWnd As Long
        Dim mhDC As Long
    
        tTop = overlay.Top
        tLeft = overlay.Left
        tWidth = overlay.Right - tLeft
        tHeight = overlay.Bottom - tTop
        
        If Abs(tHeight) < Abs(tWidth) Then
            picDisplay.FontSize = (Abs(tHeight) / 3) / zfact
        Else
            picDisplay.FontSize = (Abs(tWidth) / 3) / zfact
        End If
    
        picDisplay.FontName = "Arial"
        
        With overlay
            xc = CLng(.Left + (.Right - .Left) / 2) '/ zfact - HScroll1.Value + offsetX
            yc = CLng(.Top + (.Bottom - .Top) / 2) '/ zfact - VScroll1.Value + offsetY
        End With
         
        ImageToScreenPt xc, yc
        
        Dim hFont As Long
        Dim prevFont As Long
        
        Dim newFont As New CFont
        
        newFont.CreateFont "Arial", picDisplay.FontSize, mRotation - overlay.Rotation
        
        Debug.Print mRotation - overlay.Rotation
        
        prevFont = SelectObject(picDisplay.hdc, newFont.Handle)
 
        Dim tw As Long
        Dim th As Long
 
        Select Case mRotation - overlay.Rotation
        Case 270, -90
            th = (picDisplay.TextWidth(sTemp) / Screen.TwipsPerPixelX) / 2
            tw = -(picDisplay.TextHeight(sTemp) / Screen.TwipsPerPixelY) / 2
        
        Case 90, -270
            th = -(picDisplay.TextWidth(sTemp) / Screen.TwipsPerPixelX) / 2
            tw = (picDisplay.TextHeight(sTemp) / Screen.TwipsPerPixelY) / 2
        
        Case 180, -180
            th = -(picDisplay.TextWidth(sTemp) / Screen.TwipsPerPixelX) / 2
            tw = -(picDisplay.TextHeight(sTemp) / Screen.TwipsPerPixelY) / 2
        
        Case Else
            tw = (picDisplay.TextWidth(sTemp) / Screen.TwipsPerPixelX) / 2
            th = (picDisplay.TextHeight(sTemp) / Screen.TwipsPerPixelY) / 2
        End Select
        
        
        mTOut = TextOut(picDisplay.hdc, xc - tw, yc - th, sTemp, Len(sTemp))
        
        SelectObject picDisplay.hdc, prevFont
        
        newFont.Dispose
        Set newFont = Nothing
    End If
End Sub

Private Sub DrawAnnotations()
    Dim overlay As IOverlay
    Dim ctr As Long
    
    If mCAnnotations.Count > 0 Then
       For Each overlay In mCAnnotations
            overlay.Render Me
            NumberAnnotation overlay, ctr
            ctr = ctr + 1
            RaiseEvent AfterAnnotationDraw(picDisplay.hdc, ctr, overlay.Left, overlay.Top, overlay.Right, overlay.Bottom, overlay.Rotation)
       Next
    End If
End Sub

Public Sub Display()
    Dim picWidth As Long
    Dim picHeight As Long
    
    picWidth = picDisplay.Width / TPPX
    picHeight = picDisplay.Height / TPPY
    
    picDisplay.Cls
    
    If hImage <> 0 Then
        BitBlt picDisplay.hdc, 0, 0, picWidth, picHeight, mBuffer.hdc, 0, 0, SRCCOPY
        DrawAnnotations
        RaiseEvent AfterDisplay(picDisplay.hdc)
    End If
    Debug.Print "Display"
    picDisplay.Refresh
End Sub

' Add an annotation to the array
Public Sub AddAnnotation(ByVal lLeft As Long, ByVal lTop As Long, ByVal lWidth As Long, ByVal lHeight As Long, ByVal nRotation As Long)
    Err.Raise vbObjectError + 10, "AddAnnotation", "Relative annotating is not implemented yet"
End Sub

Public Function AddAbsoluteAnnotation(ByVal lLeft As Long, ByVal lTop As Long, ByVal lWidth As Long, ByVal lHeight As Long, ByVal nRotation As Long, ByVal lPage As Long, ByVal AnnotationType As ANNOTATION_TYPE, _
                                    ByVal FillColor As Long, ByVal LineColor As Long, ByVal Filled As Boolean, Optional Events As Boolean = True) As Long
    Dim lRect As New CRect
    Dim lCancel As Boolean
    
    Dim tempPage As Long
    Dim mTempWidth As Long
    Dim mTempHeight As Long
    
    Dim tempImage As Long
    Dim tAnnotationType As ANNOTATION_TYPE
    
    If Events Then
        RaiseEvent BeforeAnnotationCreate(lLeft, lTop, lWidth, lHeight, lCancel, AnnotationType, FillColor, LineColor, Filled)
        If lCancel Then Exit Function
    End If
    
    lRect.CreateRect2 lTop, lLeft, lWidth, lHeight

    If lRect.Left > mAbsImageWidth Or lRect.Top > mAbsImageHeight Then
        Exit Function
    End If

    ' Ignore "accidental" annotations
    If (lRect.Width > 3) And (lRect.Height > 3) Then
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
                .Page = lPage
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
                .Page = lPage
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
                .Page = lPage
            End With
        
            Set newAnnotation = newC
        
        End Select
        
        ' not on the current page!
        If lPage - 1 <> lCurPage Then
            ' grab page dimensions
            
            If hMultiBitmap <> 0 Then
                tempImage = FreeImage_LockPage(hMultiBitmap, lPage - 1)
                mBPP = FreeImage_GetBPP(tempImage)
                mTempWidth = FreeImage_GetWidth(tempImage)
                mTempHeight = FreeImage_GetHeight(tempImage)
                FreeImage_UnlockPage hMultiBitmap, tempImage, CLng(False)
            End If
            
            ' create a copy of the coordinates
            Dim tCoords As CRect
            Set tCoords = newAnnotation.GetRect
            
            ' adjust accordingly
            With newAnnotation
                Select Case newAnnotation.Rotation
                Case 90, -270
                    .Left = tCoords.Top
                    .Top = mTempHeight - tCoords.Right
                    .Bottom = .Top + (tCoords.Width)
                    .Right = .Left + (tCoords.Height)
                
                Case -90, 270
                    .Left = mTempWidth - tCoords.Bottom
                    .Top = tCoords.Left
                    .Bottom = .Top + (tCoords.Width)
                    .Right = .Left + (tCoords.Height)
                
                Case 180, -180
                    .Left = mTempWidth - tCoords.Left
                    .Top = mTempHeight - tCoords.Top
                    .Bottom = .Top + (tCoords.Height)
                    .Right = .Left + (tCoords.Width)
                
                End Select
            End With
            Set tCoords = Nothing
            
        End If
        
        With newAnnotation
            If .Left < 0 Then .Left = 0
            If .Top < 0 Then .Top = 0
            
            If .Right < 0 Then .Right = 0
            If .Bottom < 0 Then .Bottom = 0
            
            If .Left > mAbsImageWidth Then .Left = mAbsImageWidth
            If .Top > mAbsImageHeight Then .Top = mAbsImageHeight
        
            If .Right > mAbsImageWidth Then .Right = mAbsImageWidth
            If .Bottom > mAbsImageHeight Then .Bottom = mAbsImageHeight
        
        End With
        
        mCAnnotations.Add newAnnotation
        
        If Events Then RaiseEvent AfterAnnotationCreate(mCAnnotations.Count)
        
    End If

    Set lRect = Nothing
    
    AddAbsoluteAnnotation = mCAnnotations.Count

End Function

Public Sub ClearAnnotations()
    mCSelection.Clear
    mCAnnotations.Clear
End Sub

' *** Crop image to rectangle ***
Public Sub Crop(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Dim newimage As Long
    newimage = FreeImage_Copy(hImage, Left, Top, Right, Bottom)
    UpdateImgBuffer newimage
End Sub

Public Sub CropAnnotation(Index)
Dim hCropped As Long
    
    'With mAnnotations(index).coords
    '    If hCropped <> 0 Then FreeImage_Unload (hCropped)
    '    hCropped = FreeImage_Copy(hImage, .Left, .Top, .Right, .Bottom)
    '    UpdateImgBuffer hCropped
    'End With
End Sub

Public Sub FitTo(Mode As eFitTo)
    If hImage <> 0 Then
        Select Case Mode
        Case FIT_ACTUAL_PIXELS
            zfact = 1

        Case FIT_WIDTH
            zfact = (mImageWidth / (picDisplay.Width / TPPX))

        Case FIT_HEIGHT
            zfact = (mImageHeight / (picDisplay.Height / TPPY))

        Case FIT_BEST
            If picDisplay.Width / mImageWidth < picDisplay.Height / mImageHeight Then
                zfact = (mImageWidth / (picDisplay.Width / TPPX))
            Else
                zfact = (mImageHeight / (picDisplay.Height / TPPY))
            End If

        End Select
        UserControl_Resize
    End If
End Sub

Public Sub RemoveAnnotation(ByVal lIndex As Long, Optional ByVal ReflectToFreeThumb As Boolean = True)
Dim i As Long
Dim ctr As Long
Dim overlay As Object
Dim overlayobject As IOverlay

    If lIndex > -1 Then
        
        If mCAnnotations.Count > 0 Then
           For Each overlay In mCAnnotations
                Set overlayobject = overlay
                
                If ctr = lIndex Then
                    mCSelection.Remove overlayobject
                    mCAnnotations.Remove overlayobject
                    Set overlayobject = Nothing
                    Exit For
                End If

                ctr = ctr + 1
           Next
           Debug.Print "RemoveAnnotation"
           Display
        End If
            
        RaiseEvent AnnotationRemove(lIndex)

        lSelected = -1

        picDisplay.MousePointer = vbDefault
    
    End If

End Sub

' Make sure our annotations keep up with our rotations
' DEPRECATED
Private Sub RotateAnnotations(Angle)
'Dim I As Long
'
'    For I = 0 To mAnnotationCount - 1
'        mAnnotations(I).coords = RotateRect(mAnnotations(I).coords, Angle)
'    Next
End Sub

' *** Image Rotation Functions ***
Public Sub Flip()
    Dim newimage As Long
    newimage = FreeImage_RotateClassic(hImage, 180)
    mRotation = mRotation + 180
    UpdateImgBuffer newimage
    'RotateAnnotations 180
    If mRotation > 360 Then mRotation = mRotation - 360
End Sub

Public Sub RotateLeft()
    Dim newimage As Long
    newimage = FreeImage_RotateClassic(hImage, 90)
    mRotation = mRotation + 90
    mRotateDirection = Rotate_Left
    If mRotation > 360 Then mRotation = mRotation - 360
    UpdateImgBuffer newimage
End Sub

Public Sub RotateRight()
    Dim newimage As Long
    newimage = FreeImage_RotateClassic(hImage, -90)
    mRotation = mRotation - 90
    mRotateDirection = Rotate_Right
    If mRotation < 0 Then mRotation = 360 + mRotation
    UpdateImgBuffer newimage
End Sub

Public Sub Invert()
    Dim newimage As Long
    FreeImage_Invert hImage
End Sub

Public Sub Refresh()
    Debug.Print "Refresh"
    UserControl_Resize
    Display
End Sub

' "Rotates" the corners of a RECT by a specified angle
Friend Function RotateRect(tRect As CRect, ByVal Angle) As CRect
    Dim nRect As New CRect

    With nRect
        Select Case Angle
        Case 90, -270
            .Left = tRect.Top
            .Top = mAbsImageWidth - tRect.Right
            .Right = tRect.Bottom
            .Bottom = mAbsImageWidth - tRect.Left

        Case -90, 270
            .Left = mAbsImageHeight - tRect.Bottom
            .Top = tRect.Left
            .Right = mAbsImageHeight - tRect.Top
            .Bottom = tRect.Right

        Case 180, -180
            .Left = mAbsImageWidth - tRect.Right
            .Top = mAbsImageHeight - tRect.Bottom
            .Right = mAbsImageWidth - tRect.Left
            .Bottom = mAbsImageHeight - tRect.Top

        Case 0, 360
            .Left = tRect.Left
            .Top = tRect.Top
            .Right = tRect.Right
            .Bottom = tRect.Bottom

        End Select
    End With

    Set RotateRect = nRect
End Function

Friend Function UnRotateRect(tRect As CRect, ByVal Angle) As CRect
    Set UnRotateRect = New CRect

    With UnRotateRect
        Select Case Angle
        Case 90, -270
            .Left = mAbsImageWidth - tRect.Bottom
            .Top = tRect.Left
            .Right = mAbsImageWidth - tRect.Top
            .Bottom = tRect.Right

        Case -90, 270
            .Left = tRect.Top
            .Top = mAbsImageHeight - tRect.Right
            .Right = tRect.Bottom
            .Bottom = mAbsImageHeight - tRect.Left

        Case 180, -180
            .Left = mAbsImageWidth - tRect.Right
            .Top = mAbsImageHeight - tRect.Bottom
            .Right = mAbsImageWidth - tRect.Left
            .Bottom = mAbsImageHeight - tRect.Top

        Case 0, 360
            .Left = tRect.Left
            .Top = tRect.Top
            .Right = tRect.Right
            .Bottom = tRect.Bottom

        End Select
    End With

End Function

Private Sub UnRotatePoint(X, Y, Angle)
    Dim tx
    Dim ty

    Select Case Angle
    Case 90, -270
        tx = mAbsImageWidth - Y
        ty = X

    Case -90, 270
        tx = Y
        ty = mAbsImageHeight - X

    Case 180, -180
        tx = mAbsImageWidth - X
        ty = mAbsImageHeight - Y

    Case 0, 360
        tx = X
        ty = Y

    End Select

    X = tx
    Y = ty
End Sub

Private Sub RotatePoint(X, Y, Angle)
    Dim tx
    Dim ty

    Select Case Angle
    Case 90, -270
        tx = Y
        ty = mAbsImageWidth - X

    Case -90, 270
        tx = mAbsImageHeight - Y
        ty = X

    Case 180, -180
        tx = mAbsImageWidth - X
        ty = mAbsImageHeight - Y

    Case 0, 360
        tx = X
        ty = Y

    End Select

    X = tx
    Y = ty
End Sub

' *** Updates the in-memory image after image-destroying operations such as rotate
Private Sub UpdateImgBuffer(ByVal hNewHandle As Long)
    If hNewHandle <> 0 Then
        If hImage <> 0 Then FreeImage_Unload hImage

        ' point hImage to new image
        hImage = hNewHandle

        ' update size variables
        mBPP = FreeImage_GetBPP(hImage)
        mImageWidth = FreeImage_GetWidth(hImage)
        mImageHeight = FreeImage_GetHeight(hImage)
                
        'Refresh
        UserControl_Resize
    End If
End Sub

Private Sub RotateImage(ImageAnnotationRotation As Long)
Dim newimage As Long
Dim cRotation As Integer
Dim SubT As Integer
        
    ImageAnnotationRotation = IIf(ImageAnnotationRotation = 360, 0, ImageAnnotationRotation)
    mRotation = IIf(mRotation = 360, 0, mRotation)
    
    SubT = ImageAnnotationRotation - mRotation
    cRotation = SubT
    newimage = FreeImage_RotateClassic(hImage, cRotation)
    
    mRotation = ImageAnnotationRotation
    UpdateImgBufferRotatedWithZoom newimage
End Sub

Private Sub UpdateImgBufferRotatedWithZoom(ByVal hNewHandle As Long)
    If hNewHandle <> 0 Then
        If hImage <> 0 Then FreeImage_Unload hImage
        ' point hImage to new image
        hImage = hNewHandle
        ' update size variables
        mBPP = FreeImage_GetBPP(hImage)
        mImageWidth = FreeImage_GetWidth(hImage)
        mImageHeight = FreeImage_GetHeight(hImage)
                
        'Refresh
        'UserControl_Resize
    End If
End Sub

Public Sub ClearMemory()
    If hCopy <> 0 Then
        FreeImage_Unload hCopy
        hCopy = 0
    End If
End Sub

Public Sub LoadFromMemory()
    Dim newimage As Long
    If hCopy <> 0 Then
        If hImage <> 0 Then
            FreeImage_Unload hImage
            hImage = 0
        End If

        mBPP = FreeImage_GetBPP(hCopy)
        mImageWidth = FreeImage_GetWidth(hCopy)
        mImageHeight = FreeImage_GetHeight(hCopy)

        newimage = FreeImage_Copy(hCopy, 0, 0, mImageWidth, mImageHeight)
        UpdateImgBuffer newimage

        mAbsImageWidth = mImageWidth
        mAbsImageHeight = mImageHeight

        mRotation = 0

    End If
End Sub

Public Sub SaveToMemory()
    If hCopy <> 0 Then
        FreeImage_Unload hCopy
        hCopy = 0
    End If

    hCopy = FreeImage_Copy(hImage, 0, 0, mImageWidth, mImageHeight)
End Sub

Private Sub picDisplay_Click()
    RaiseEvent Click
End Sub

Private Sub picDisplay_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub picDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub picDisplay_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub picDisplay_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub picDisplay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim DeltaX As Single
    Dim DeltaY As Single
    Dim overlay As IOverlay
    
    If Not bLocked Then
        
        ' Check which button is being pressed
        Select Case Button
        Case vbLeftButton
            
            If lAnnotate = 1 Then lAnnotate = 2
            
            If bMouseScroll Then           ' bMouseScroll
                ' Scroll the image by dragging the mouse
                bUpdate = False

                DeltaX = (dx - X) / TPPX
                DeltaY = (dy - Y) / TPPY

                If VScroll1.Max > 0 Then
                    If VScroll1.Value + DeltaY < 0 Then
                        VScroll1.Value = 0
                    ElseIf VScroll1.Value + DeltaY > VScroll1.Max Then
                        VScroll1.Value = VScroll1.Max
                    Else
                        VScroll1.Value = VScroll1.Value + DeltaY
                        pY = pY - dy + Y
                    End If
                End If

                If HScroll1.Max > 0 Then
                    If HScroll1.Value + DeltaX < 0 Then
                        HScroll1.Value = 0
                    ElseIf HScroll1.Value + DeltaX > HScroll1.Max Then
                        HScroll1.Value = HScroll1.Max
                    Else
                        HScroll1.Value = HScroll1.Value + DeltaX
                        pX = pX - dx + X
                    End If
                End If

                bUpdate = True

                dx = X
                dy = Y

                ScrollDisplay

            Else                         'bMouseScroll
                
                If mCSelection.Count > 0 Then   'lSelected > -1

                    If bFastSelection Then
                        If bChanged Then DrawSelection
                    End If
                    
                    If mCAnnotations.Count > 0 Then
                    
                        For Each overlay In mCSelection
                             
                             Dim pdx As Single
                             Dim pdy As Single
                             
                             pdx = (X - dx) / TPPX * zfact
                             pdy = (Y - dy) / TPPY * zfact
                             
                             'If Abs(pdx) > 0 And Abs(pdx) < 1 Then pdx = Sgn(pdx) / TPPX * zfact
                             'If Abs(pdy) > 0 And Abs(pdy) < 1 Then pdy = Sgn(pdy) / TPPX * zfact
                             
                             'Debug.Print pdx, pdy
                             
                             If overlay.Selected Then
                                 overlay.Move pdx, pdy, Me
                             End If
                        Next
                        
                        'Display
                    End If

                    dx = X
                    dy = Y

                    bChanged = True

                    If bFastSelection Then
                        DrawSelection
                    Else
                        Debug.Print "MouseMove"
                        Display
                    End If
                    
                Else                     'lSelected <= -1

                    DrawRubberBand
                    
                    If X > picDisplay.Width Then
                        If HScroll1.Value + (X - picDisplay.Width) / 10 > HScroll1.Max Then
                            HScroll1.Value = HScroll1.Max
                        Else
                            HScroll1.Value = HScroll1.Value + (X - picDisplay.Width) / 10
                            pX = pX - dH * TPPX
                        End If
                    End If

                    If Y > picDisplay.Height Then
                        If VScroll1.Value + (Y - picDisplay.Height) / 10 > VScroll1.Max Then
                            VScroll1.Value = VScroll1.Max
                        Else
                            VScroll1.Value = VScroll1.Value + (Y - picDisplay.Height) / 10
                            pY = pY - dV * TPPY
                        End If
                    End If

                    If X <= 0 Then
                        If HScroll1.Value + X / 10 < 0 Then
                            HScroll1.Value = 0
                        Else
                            HScroll1.Value = HScroll1.Value + X / 10
                            pX = pX - dH * TPPX
                        End If
                    End If

                    If Y <= 0 Then
                        If VScroll1.Value + Y / 10 < 0 Then
                            VScroll1.Value = 0
                        Else
                            VScroll1.Value = VScroll1.Value + Y / 10
                            pY = pY - dV * TPPY
                        End If
                    End If


                    dx = X
                    dy = Y

                    bPreShowSelect = True

                    DrawRubberBand
                    
                    If X - offsetX * TPPX < 0 Then X = offsetX * TPPX
                    If Y - offsetY * TPPY < 0 Then Y = offsetY * TPPY
                    If X - offsetX * TPPX > mImageWidth / zfact * TPPX Then X = offsetX * TPPX + mImageWidth / zfact * TPPX
                    If Y - offsetY * TPPY > mImageHeight / zfact * TPPX Then Y = offsetY * TPPY + mImageHeight / zfact * TPPY
                
                End If                   'lSelected > -1

            End If                       'bMouseScroll

        Case 0  ' no mousebutton pressed
            
                'If mCSelection.Count > 0 Then
                   For Each overlay In mCAnnotations
                        Dim hitType As HitTestEnum
                        
                        hitType = overlay.HitTest(X / TPPX, Y / TPPY, Me)
                        
                        Select Case hitType
                        Case HitTestEnum.HIT_MOVE, HitTestEnum.HIT_START, HitTestEnum.HIT_END
                            UserControl.MousePointer = vbSizeAll
                            Exit For
                            
                        Case HitTestEnum.HIT_NONE, HitTestEnum.HIT_CENTER
                            UserControl.MousePointer = vbDefault
                            'Exit For    ' avoid selecting below
                        
                        Case HitTestEnum.HIT_TOP, HitTestEnum.HIT_BOTTOM
                            UserControl.MousePointer = vbSizeNS
                            Exit For    ' avoid selecting below
                        
                        Case HitTestEnum.HIT_LEFT, HitTestEnum.HIT_RIGHT
                            UserControl.MousePointer = vbSizeWE
                            Exit For    ' avoid selecting below
                        
                        Case HitTestEnum.HIT_LEFT + HitTestEnum.HIT_TOP, HitTestEnum.HIT_RIGHT + HitTestEnum.HIT_BOTTOM
                            UserControl.MousePointer = vbSizeNWSE
                            Exit For    ' avoid selecting below
                        
                        Case HitTestEnum.HIT_RIGHT + HitTestEnum.HIT_TOP, HitTestEnum.HIT_LEFT + HitTestEnum.HIT_BOTTOM
                            UserControl.MousePointer = vbSizeNESW
                            Exit For    ' avoid selecting below
                        End Select
                   Next
                'Else
                '    UserControl.MousePointer = vbDefault
                'End If
                
        End Select

    End If

    mCursorX = (X / TPPX - offsetX + CLng(HScroll1.Value)) * zfact
    mCursorY = (Y / TPPY - offsetY + CLng(VScroll1.Value)) * zfact

    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub DrawRubberBand()
    If bPreShowSelect Then
        picDisplay.DrawMode = vbInvert
        picDisplay.DrawStyle = vbDot
        Select Case mRubberBand
        Case RUBBERBAND_BOX
            picDisplay.Line (pX, pY)-(dx, dy), , B
        Case RUBBERBAND_LINE
            picDisplay.Line (pX, pY)-(dx, dy)
        End Select
    End If
End Sub

Private Sub picDisplay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim tx As Single, ty As Single

    RaiseEvent MouseUp(Button, Shift, X, Y)

    If Not bLocked Then
        
        bEdit = False
        bMove = False
        bPreShowSelect = False
            
        If mCSelection.Count = 0 Then
            
            'If Not bMouseScroll Then
            
            If lAnnotate = 2 Then
                
                If X - offsetX * TPPX < 0 Then X = offsetX * TPPX
                If Y - offsetY * TPPY < 0 Then Y = offsetY * TPPY
                If X - offsetX * TPPX > mImageWidth / zfact * TPPX Then X = (offsetX + mImageWidth / zfact) * TPPX
                If Y - offsetY * TPPY > mImageHeight / zfact * TPPY Then Y = (offsetY + mImageHeight / zfact) * TPPY
                
                Dim pRect As CRect
                Set pRect = ScreenToImage(pX, pY, X, Y)
                
                If mAutoAnnotate Then
                    AddAbsoluteAnnotation pRect.Left, pRect.Top, pRect.Width, pRect.Height, mRotation, lCurPage + 1, mDefaultAnnotationType, mDefaultAnnotationFillColor, mDefaultAnnotationLineColor, mAnnotationFilled
                Else
                    RaiseEvent AnnotationCreate(pRect.Left, pRect.Top, pRect.Width, pRect.Height, mRotation, lCurPage + 1)
                End If
                
                Set pRect = Nothing
                
                Debug.Print "MouseUp"
                Display
        
                lAnnotate = 0
            
            End If  ' lAnnotate = 2
            
            'End If
            
        Else ' mCSelection.Count = 0
            
            ' Finished editing an annotation
            mCAnnotations.CallUpdate
        
        End If ' mCSelection.Count = 0
    
    End If ' not bLocked

End Sub

Private Sub picDisplay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim tRect As RECT
Dim pRect As RECT
Dim tx As Long, ty As Long
Dim i As Long
Dim j As Long

    dx = X
    dy = Y

    pX = X
    pY = Y

    tx = (X / TPPX)
    ty = (Y / TPPY)
    
    txDown = tx
    tyDown = ty
    
    lAnnotate = 1

    Dim overlay As Variant
    Dim overlayobject As IOverlay
    Dim ctr As Long

    For Each overlay In mCSelection
         Set overlayobject = overlay
         overlayobject.Selected = False
    Next
       
    mCSelection.Clear
    
    lSelected = -1
    
    If mCAnnotations.Count > 0 Then
       For Each overlay In mCAnnotations
            Set overlayobject = overlay
            
            Dim hitType As HitTestEnum
            hitType = overlayobject.HitTest(X / TPPX, Y / TPPY, Me)
            
            If hitType <> HIT_NONE Then
                overlayobject.Selected = True
                If hitType = HIT_CENTER Then
                    RaiseEvent AnnotationClick(ctr)
                End If
                lSelected = ctr
                mCSelection.Add overlayobject
                Exit For    ' avoid selecting below
            End If
            ctr = ctr + 1
       Next
       
       Debug.Print "MouseDown"
       Display
    
    End If

    If X - offsetX * TPPX < 0 Then pX = offsetX * TPPX
    If Y - offsetY * TPPY < 0 Then pY = offsetY * TPPY
    If X - offsetX * TPPX > mImageWidth / zfact * TPPX Then pX = offsetX * TPPX + mImageWidth / zfact * TPPX
    If Y - offsetY * TPPY > mImageHeight / zfact * TPPY Then pY = offsetY * TPPY + mImageHeight / zfact * TPPY

    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub HScroll1_Change()
    dH = HScroll1.Value - lastHScroll
    dV = 0
    
    lastHScroll = HScroll1.Value
    ScrollDisplay
    
    HandleScrollEvents
End Sub

Private Sub HScroll1_Scroll()
    dH = HScroll1.Value - lastHScroll
    dV = 0
    
    lastHScroll = HScroll1.Value
    ScrollDisplay
    
    HandleScrollEvents
End Sub

Private Sub VScroll1_Change()
    dV = VScroll1.Value - lastVScroll
    dH = 0
    
    lastVScroll = VScroll1.Value
    ScrollDisplay
    
    HandleScrollEvents
End Sub

Private Sub VScroll1_Scroll()
    dV = VScroll1.Value - lastVScroll
    dH = 0

    lastVScroll = VScroll1.Value
    ScrollDisplay
    
    HandleScrollEvents
End Sub

Private Sub HandleScrollEvents()
    If bScrollEvents Then
        'Ely ScrollChange
        UpdateWindowEvent
        
        RaiseEvent ScrollChange(HScroll1.Value * zfact, VScroll1.Value * zfact)
    End If
End Sub

Private Sub ScrollDisplay()
Dim picWidth As Long
Dim picHeight As Long
Dim hBrush As Long
    
    picWidth = picDisplay.Width / TPPX
    picHeight = picDisplay.Height / TPPY
    
    Debug.Print dH, dV
    
    If dH > 0 Then
        BitBlt mBuffer.hdc, 0, 0, picWidth - dH, picHeight, mBuffer.hdc, dH, 0, SRCCOPY
        FreeImage_PaintDCEx mBuffer.hdc, hImage, picWidth - dH, -offsetY, dH, picHeight, (HScroll1.Value + picWidth - dH) * zfact, (VScroll1.Max - VScroll1.Value) * zfact, dH * zfact, picHeight * zfact, DM_DRAW_DEFAULT, ROP_SRCCOPY
    End If

    If dH < 0 Then
        BitBlt mBuffer.hdc, -dH, 0, picWidth + dH, picHeight, mBuffer.hdc, 0, 0, SRCCOPY
        FreeImage_PaintDCEx mBuffer.hdc, hImage, 0, -offsetY, Abs(dH), picHeight, HScroll1.Value * zfact, (VScroll1.Max - VScroll1.Value) * zfact, Abs(dH) * zfact, picHeight * zfact, DM_DRAW_DEFAULT, ROP_SRCCOPY
    End If

    If dV > 0 Then
        BitBlt mBuffer.hdc, 0, 0, picWidth, picHeight - dV, mBuffer.hdc, 0, dV, SRCCOPY
        FreeImage_PaintDCEx mBuffer.hdc, hImage, offsetX, picHeight - dV, picWidth, dV, (HScroll1.Value) * zfact, (VScroll1.Max - VScroll1.Value) * zfact, picWidth * zfact, dV * zfact, DM_DRAW_DEFAULT, ROP_SRCCOPY
    End If

    If dV < 0 Then
        BitBlt mBuffer.hdc, 0, -dV, picWidth, picHeight + dV, mBuffer.hdc, 0, 0, SRCCOPY
        FreeImage_PaintDCEx mBuffer.hdc, hImage, offsetX, 0, picWidth, Abs(dV), (HScroll1.Value) * zfact, (VScroll1.Max - VScroll1.Value + picHeight + dV) * zfact, picWidth * zfact, Abs(dV) * zfact, DM_DRAW_DEFAULT, ROP_SRCCOPY
    End If
    
    Debug.Print "ScrollDisplay"
    If bUpdate Then Display
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Locked", bLocked
    PropBag.WriteProperty "ScrollBarsVisible", bShowScrollBars
    PropBag.WriteProperty "SelectedColor", mSelectedColor
    PropBag.WriteProperty "BackColor", picDisplay.BackColor
    PropBag.WriteProperty "BackgroundImage", picDisplay.Picture
    PropBag.WriteProperty "LineWidth", mLineWidth
    PropBag.WriteProperty "AnnotationFilled", mAnnotationFilled
    PropBag.WriteProperty "ShowNumber", mShowNumber
    
    PropBag.WriteProperty "DefaultAnnotationType", mDefaultAnnotationType
    PropBag.WriteProperty "DefaultAnnotationFillColor", mDefaultAnnotationFillColor
    PropBag.WriteProperty "DefaultAnnotationLineColor", mDefaultAnnotationLineColor
    PropBag.WriteProperty "ReflectRotationOnAnnotation", mReflectRotationOnAnnotation
    'PropBag.WriteProperty "FreeThumbEdit", MFreeThumb
    PropBag.WriteProperty "Rubberband", mRubberBand

End Sub

Public Sub GetAbsoluteAnnotation(ByVal Index As Long, Left As Long, Top As Long, Width As Long, Height As Long, Rotation As Long, Page As Long)
    Dim tempRect As CRect
    Dim overlay As IOverlay

    Set overlay = mCAnnotations(Index + 1)

    With overlay
        Left = .Left
        Top = .Top
        Width = .Right - .Left
        Height = .Bottom - .Top
        Rotation = .Rotation
        Page = .Page + 1
    End With

End Sub

Public Sub GetAnnotation(ByVal Index As Long, Left As Long, Top As Long, Width As Long, Height As Long, Rotation As Long)
    Dim thisAnnotation As CAnnotation
    Set thisAnnotation = mCAnnotations(Index + 1)
    
    If Not thisAnnotation Is Nothing Then
        With thisAnnotation
            Left = .Left
            Top = .Top
            Width = .Width
            Height = .Height
            Rotation = .Rotation
        End With
    End If

End Sub

Public Sub SetAnnotation(ByVal Index As Long, Left As Long, Top As Long, Width As Long, Height As Long, Rotation As Long)
    Dim thisAnnotation As CAnnotation
    Set thisAnnotation = mCAnnotations(Index + 1)
    
    If Not thisAnnotation Is Nothing Then
        With thisAnnotation
            .Left = Left
            .Top = Top
            .Width = Width
            .Height = Height
            .Rotation = Rotation
        End With
    End If
    
End Sub

Public Sub MoveAnnotation(ByVal Index As Long, ByVal DeltaX As Long, ByVal DeltaY As Long)
    Dim dx As Long
    Dim dy As Long

    If Index > -1 Then

        ' switch deltax and deltay around depending on the current rotation
        Select Case mRotation
        Case 0, 360
            dx = DeltaX
            dy = DeltaY

        Case 90, -270
            dx = -DeltaY
            dy = DeltaX

        Case 180, -180
            dx = -DeltaX
            dy = -DeltaY

        Case -90, 270
            dx = DeltaY
            dy = -DeltaX

        End Select

'        With mAnnotations(Index).coords
'            ' better dragging behaviour at image edges
'            If .Left + dx < 0 Then dx = -.Left
'            If .Right + dx > mAbsImageWidth Then dx = mAbsImageWidth - .Right
'            If .Top + dy < 0 Then dy = -.Top
'            If .Bottom + dy > mAbsImageHeight Then dy = mAbsImageHeight - .Bottom
'
'            .Left = .Left + dx
'            .Top = .Top + dy
'            .Right = .Right + dx
'            .Bottom = .Bottom + dy
'
'        End With

    End If
End Sub

Private Sub MovePoint(Optional mLeft, Optional mTop, Optional mRight, Optional mBottom, Optional DeltaX As Long, Optional DeltaY As Long)
    Dim dx As Long
    Dim dy As Long

    ' switch deltax and deltay around depending on the current rotation
    Select Case mRotation
    Case 0, 360
        dx = DeltaX
        dy = DeltaY

    Case 90, -270
        dx = -DeltaY
        dy = DeltaX

    Case 180, -180
        dx = -DeltaX
        dy = -DeltaY

    Case -90, 270
        dx = DeltaY
        dy = -DeltaX

    End Select

    ' better dragging behaviour at image edges
    If Not IsMissing(mLeft) Then
        If mLeft + dx < 0 Then dx = -mLeft
        mLeft = mLeft + dx
    End If

    If Not IsMissing(mTop) Then
        If mTop + dy < 0 Then dy = -mTop
        mTop = mTop + dy
    End If

    If Not IsMissing(mRight) Then
        If mRight + dx > mAbsImageWidth Then dx = mAbsImageWidth - mRight
        mRight = mRight + dx
    End If

    If Not IsMissing(mBottom) Then
        If mBottom + dy > mAbsImageHeight Then dy = mAbsImageHeight - mBottom
        mBottom = mBottom + dy
    End If
End Sub

'Public Sub SetAnnotation(ByVal Index As Long, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long)
'On Local Error GoTo errhand
'    With mAnnotations(Index).coords
'        .Left = Left
'        .Top = Top
'        .Right = Left + Width
'        .Bottom = Top + Height
'    End With
'errhand:
'    If Err.Number = 9 Then
'        Debug.Print Err.Number
'    End If
'End Sub

Public Sub EditAnnotation(ByVal Index As Long, ByVal HitPoint As HitTestEnum, ByVal rX As Long, ByVal rY As Long)
Dim thisAnnotation As IOverlay
    Set thisAnnotation = mCAnnotations(Index)
    thisAnnotation.HitPoint = HitPoint
    thisAnnotation.Move rX, rY, Me
End Sub


' Returns TRUE if an image contains less than (threshold) white pixels
' Requires Weights() to be filled in with appropriate values
Public Function IsInverted(Threshold As Integer) As Boolean
    Dim ImageData() As Byte
    Dim lLine As Long
    Dim lPitch As Long
    Dim lWhite As Long
    Dim lBlack As Long
    Dim lBits As Long
    Dim Y As Long
    Dim X As Long
    

    If hImage <> 0 Then
        If FreeImage_GetBPP(hImage) = 1 Then
            lLine = FreeImage_GetLine(hImage) * 8
            For Y = 0 To FreeImage_GetHeight(hImage)
                ImageData = FreeImage_GetScanLineEx(hImage, Y)
                lBits = 0
                For X = 0 To UBound(ImageData)
                    lBits = lBits + Weights(ImageData(X))
                Next
                lBlack = lBlack + lBits
                lWhite = lWhite + lLine - lBits

                Call FreeImage_DestroyLockedArray(ImageData)
            Next
            lBlack = lBlack + 1
            IsInverted = (lWhite / lBlack) * 100 < Threshold

        Else
            Err.Raise vbObjectError + 10, "IsInverted", "Image is not a bilevel"

        End If
    End If
End Function

Private Function ImageToScreen(tRect As CRect) As CRect
    Set ImageToScreen = RotateRect(tRect, mRotation)

    ImageToScreen.Top = offsetY + (ImageToScreen.Top) / zfact - VScroll1.Value
    ImageToScreen.Left = offsetX + (ImageToScreen.Left) / zfact - HScroll1.Value
    ImageToScreen.Right = offsetX + (ImageToScreen.Right) / zfact - HScroll1.Value
    ImageToScreen.Bottom = offsetY + (ImageToScreen.Bottom) / zfact - VScroll1.Value
End Function

Private Function ScreenToImage(x1, y1, x2, y2) As CRect
    Dim tempRect As CRect
    Set tempRect = New CRect
    
    tempRect.Top = (y1 / TPPY - offsetY + VScroll1.Value) * zfact
    tempRect.Left = (x1 / TPPX - offsetX + HScroll1.Value) * zfact
    tempRect.Bottom = (y2 / TPPY - offsetY + VScroll1.Value) * zfact
    tempRect.Right = (x2 / TPPX - offsetX + HScroll1.Value) * zfact
    
    Set ScreenToImage = UnRotateRect(tempRect, mRotation)
    
    Set tempRect = Nothing
End Function

Private Function ScreenToImagePt(x1, y1)
    x1 = (x1 / TPPX - offsetX + HScroll1.Value) * zfact
    y1 = (y1 / TPPY - offsetY + VScroll1.Value) * zfact

    UnRotatePoint x1, y1, mRotation
    
End Function

Friend Function ImageToScreenPt(x1, y1)
    RotatePoint x1, y1, mRotation

    x1 = x1 / zfact - HScroll1.Value + offsetX
    y1 = y1 / zfact - VScroll1.Value + offsetY
End Function

Public Function ZoomToAnnotation(Index)
    If Not Index = -1 Then
        ZoomHack Index
    End If
End Function

' FOR FIXING
Private Function ZoomHack(Index)
    Dim lWidth As Long
    Dim lHeight As Long
    Dim hVal As Long
    Dim vVal As Long
    
    Dim nRect As CRect
    Dim tRect As CRect
    Dim tRectHolderRotate As CRect
        
    Dim thisAnnotation As IOverlay
        
    Set thisAnnotation = mCAnnotations(Index + 1)
    Set nRect = thisAnnotation.GetRect()
    Set tRect = ImageToScreen(nRect)
    
    'If mReflectRotationOnAnnotation = True Then
    If Not thisAnnotation.Rotation = mRotation Then
        RotateImage thisAnnotation.Rotation
    End If
    'End If
    
    With nRect
        Select Case mRotation Mod 360
        Case 0, 180
                lWidth = .Width
                lHeight = .Height
        
        Case 270, 90
                lWidth = .Height
                lHeight = .Width
        End Select
    End With
    
    If picDisplay.Width / lWidth < picDisplay.Height / lHeight Then
        zfact = (lWidth / (picDisplay.Width / TPPX))
    Else
        zfact = (lHeight / (picDisplay.Height / TPPY))
    End If

    If zfact < 1 / 6 Then zfact = 1 / 6
    
    UserControl_Resize
    
    Set tRectHolderRotate = New CRect
    tRectHolderRotate.CopyRect nRect
    
    With tRectHolderRotate
        Select Case mRotation Mod 360
        Case 0
            vVal = offsetY + .Top / zfact
            hVal = offsetX + .Left / zfact
        
        Case 270
            tRectHolderRotate.Top = .Right
            tRectHolderRotate.Left = .Top
            tRectHolderRotate.Bottom = .Left
            tRectHolderRotate.Right = .Bottom
            
            vVal = (offsetY + tRectHolderRotate.Top / zfact) - picDisplay.Height / TPPY
            hVal = HScroll1.Max - (offsetX + tRectHolderRotate.Left / zfact)
        
        Case 180
            tRectHolderRotate.Top = .Bottom
            tRectHolderRotate.Left = .Right
            tRectHolderRotate.Bottom = .Top
            tRectHolderRotate.Right = .Left
            
            vVal = VScroll1.Max - (offsetY + tRectHolderRotate.Top / zfact) + picDisplay.Height / TPPY
            hVal = HScroll1.Max - (offsetX + tRectHolderRotate.Left / zfact) + picDisplay.Width / TPPX
        
        Case 90
            tRectHolderRotate.Top = .Left
            tRectHolderRotate.Left = .Bottom
            tRectHolderRotate.Bottom = .Right
            tRectHolderRotate.Right = .Top
        
            vVal = VScroll1.Max - (offsetY + tRectHolderRotate.Top / zfact)
            hVal = (offsetX + tRectHolderRotate.Left / zfact) - picDisplay.Width / TPPX
        
        End Select
    End With
    
    bUpdate = False
    On Error Resume Next
    If vVal < VScroll1.Max Then VScroll1.Value = vVal Else VScroll1.Value = VScroll1.Max
    If hVal < HScroll1.Max Then HScroll1.Value = hVal Else HScroll1.Value = HScroll1.Max
    On Error GoTo 0
    bUpdate = True

    Set nRect = Nothing
    Set tRect = Nothing
End Function

Public Sub ConvertColorDepth(Mode As Long)
    hImage = FreeImage_ConvertColorDepth(hImage, Mode, True)
End Sub

Public Sub ThreshHold(lThresh As Long)
    hImage = FreeImage_ConvertColorDepth(hImage, FICF_MONOCHROME_THRESHOLD, True, CByte(lThresh))
End Sub

Public Function AdjustBrightness(percentage As Double) As Boolean
    AdjustBrightness = (FreeImage_AdjustBrightness(hImage, percentage) <> 0)
End Function

Public Function AdjustContrast(percentage As Double) As Boolean
    AdjustContrast = (FreeImage_AdjustContrast(hImage, percentage) <> 0)
End Function

Public Function TextWidth(Text As String) As Long
    TextWidth = picDisplay.TextWidth(Text)
End Function

Public Function TextHeight(Text As String) As Long
    TextHeight = picDisplay.TextHeight(Text)
End Function

Property Let FontSize(v As Long)
    If v = 0 Then v = 1
    picDisplay.FontSize = v
End Property

Property Get FontSize() As Long
    FontSize = picDisplay.FontSize
End Property

Property Let FontName(v As String)
    picDisplay.FontName = v
End Property

Property Get FontName() As String
    FontName = picDisplay.FontName
End Property

Property Let FontBold(v As Boolean)
    picDisplay.FontBold = v
End Property

Property Get FontBold() As Boolean
    FontBold = picDisplay.FontBold
End Property

Property Let ForeColor(v As OLE_COLOR)
    picDisplay.ForeColor = v
End Property

Property Get ForeColor() As OLE_COLOR
    ForeColor = picDisplay.ForeColor
End Property

Property Let DrawMode(v As Long)
    picDisplay.DrawMode = v
End Property

Property Get DrawMode() As Long
    DrawMode = picDisplay.DrawMode
End Property

Property Let FastSelection(v As Boolean)
    bFastSelection = v
End Property

Property Get FastSelection() As Boolean
    FastSelection = bFastSelection
End Property

' Helper API functions (hidden)
Public Function APIBitBlt(ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, xSrc As Long, ySrc As Long, dwRop As Long) As Long
    APIBitBlt = BitBlt(picDisplay.hdc, X, Y, nWidth, nHeight, hSrcDC, xSrc, ySrc, dwRop)
End Function

Public Function APITextOut(ByVal X As Long, ByVal Y As Long, ByVal lpString As String)
    APITextOut = TextOut(picDisplay.hdc, X, Y, lpString, Len(lpString))
End Function

Public Function APIMoveTo(ByVal X As Long, ByVal Y As Long, Optional oldX As Long, Optional oldY As Long)
Dim oldPoint As POINTAPI
    APIMoveTo = MoveToEx(picDisplay.hdc, X, Y, oldPoint)
    If Not IsMissing(oldX) Then oldX = oldPoint.X
    If Not IsMissing(oldY) Then oldY = oldPoint.Y
End Function

Public Function APILineTo(ByVal X As Long, ByVal Y As Long)
    APILineTo = LineTo(picDisplay.hdc, X, Y)
End Function

Public Function APISelectObject(ByVal hObject As Long)
    APISelectObject = SelectObject(picDisplay.hdc, hObject)
End Function

Public Function APIDeleteObject(ByVal hObject As Long)
    APIDeleteObject = DeleteObject(hObject)
End Function

Friend Property Get AbsImageWidth() As Long
    AbsImageWidth = mAbsImageWidth
End Property

Friend Property Get AbsImageHeight() As Long
    AbsImageHeight = mAbsImageHeight
End Property

Public Sub Despeckle(radius As Integer, Percent As Double)
    Err.Raise vbObjectError + 201, "Despeckle", "Not implemented yet"
    ' does nothing
End Sub

Property Get Annotations() As CAnnotations
    Set Annotations = mCAnnotations
End Property

Property Set Annotations(v As CAnnotations)
    Debug.Print mCAnnotations.Name & " replaced by " & v.Name
    Set mCAnnotations = v
End Property

'Interface implementation
Private Property Get IContainer_hDC() As Long
    IContainer_hDC = hdc
End Property

Private Property Get IContainer_Height() As Long
    IContainer_Height = AbsImageHeight
End Property

Private Function IContainer_ImageToScreen(v As CRect) As CRect
    Set IContainer_ImageToScreen = ImageToScreen(v)
End Function

Private Property Get IContainer_Page() As Long
    IContainer_Page = Page
End Property

Private Property Get IContainer_Rotation() As Long
    IContainer_Rotation = Rotation
End Property

Private Property Get IContainer_Width() As Long
    IContainer_Width = AbsImageWidth
End Property

Private Property Let IContainer_DrawMode(RHS As Long)
    DrawMode = RHS
End Property

Private Property Let IContainer_ForeColor(RHS As Long)
    ForeColor = RHS
End Property

Property Set Selection(v As CAnnotations)
    Set mCSelection = v
End Property

Property Get Selection() As CAnnotations
    Set Selection = mCSelection
End Property

Property Let Rubberband(v As RubberBandMode)
    mRubberBand = v
    PropertyChanged "Rubberband"
End Property

Property Get Rubberband() As RubberBandMode
    Rubberband = mRubberBand
End Property

' Annotation Collection Events
Private Sub mCAnnotations_Updated()
    Debug.Print "mCAnnotations_Updated"
    Display
End Sub

Private Sub mCSelection_Updated()
    Debug.Print "mCSelection_Updated"
    Display
End Sub


