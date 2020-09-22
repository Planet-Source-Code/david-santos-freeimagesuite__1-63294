VERSION 5.00
Object = "*\AfreeImageSuite.vbp"
Begin VB.Form frmTest 
   Caption         =   "Test Form"
   ClientHeight    =   10065
   ClientLeft      =   885
   ClientTop       =   1485
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10065
   ScaleWidth      =   15240
   Begin FreeImageSuite.FreeThumb FreeThumb1 
      Height          =   3855
      Left            =   120
      TabIndex        =   36
      Top             =   6120
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   6800
      Rubberband      =   0
   End
   Begin FreeImageSuite.FreeImgEdit FreeImgEdit1 
      Height          =   5295
      Left            =   120
      TabIndex        =   35
      Top             =   600
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   9340
      Locked          =   0   'False
      ScrollBarsVisible=   -1  'True
      SelectedColor   =   16711680
      BackColor       =   8421504
      BackgroundImage =   "frmTest.frx":0000
      LineWidth       =   0
      AnnotationFilled=   0   'False
      ShowNumber      =   0   'False
      DefaultAnnotationType=   0
      DefaultAnnotationFillColor=   0
      DefaultAnnotationLineColor=   255
      ReflectRotationOnAnnotation=   0   'False
      Rubberband      =   0
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   9600
      TabIndex        =   34
      Text            =   "*.jpg"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Split Horizontal"
      Height          =   375
      Left            =   12480
      TabIndex        =   33
      Top             =   5280
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   32
      Text            =   "C:\Documents and Settings\Dave\My Documents\My Pictures\"
      Top             =   120
      Width           =   9375
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   14040
      Top             =   5640
   End
   Begin VB.Frame Frame3 
      Caption         =   "Thumbnail"
      Height          =   1095
      Left            =   12480
      TabIndex        =   26
      Top             =   600
      Width           =   2655
      Begin VB.TextBox txtThumbHeight 
         Height          =   285
         Left            =   1920
         TabIndex        =   29
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtThumbWidth 
         Height          =   285
         Left            =   720
         TabIndex        =   28
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdGetThumb_HeightWidth 
         Caption         =   "GET"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Height:"
         Height          =   255
         Left            =   1320
         TabIndex        =   31
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Width:"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Annotation Coordinates"
      Height          =   1815
      Left            =   12480
      TabIndex        =   11
      Top             =   2520
      Width           =   2655
      Begin VB.TextBox txtHeight 
         Height          =   285
         Left            =   2160
         TabIndex        =   24
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtWidth 
         Height          =   285
         Left            =   1080
         TabIndex        =   23
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtTop 
         Height          =   285
         Left            =   2160
         TabIndex        =   22
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtLeft 
         Height          =   285
         Left            =   1080
         TabIndex        =   21
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtAnnotationIndex 
         Height          =   285
         Left            =   2160
         TabIndex        =   14
         Text            =   "0"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtIndex 
         Height          =   285
         Left            =   2160
         TabIndex        =   13
         Text            =   "1"
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdGetCoordinates 
         Caption         =   "GET"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Height:"
         Height          =   255
         Left            =   1560
         TabIndex        =   20
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Width:"
         Height          =   255
         Left            =   600
         TabIndex        =   19
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Left:"
         Height          =   255
         Left            =   720
         TabIndex        =   18
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Top:"
         Height          =   255
         Left            =   1800
         TabIndex        =   17
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Anno Index"
         Height          =   255
         Left            =   1200
         TabIndex        =   16
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Thumb Index"
         Height          =   255
         Left            =   1200
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Annotation Count"
      Height          =   615
      Left            =   12480
      TabIndex        =   8
      Top             =   1800
      Width           =   2655
      Begin VB.CommandButton AnnotationCount 
         Caption         =   "GET"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblAnnotationCount 
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Split Vertical"
      Height          =   375
      Left            =   12480
      TabIndex        =   7
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "+"
      Height          =   375
      Left            =   14280
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "-"
      Height          =   375
      Left            =   14760
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Toggle Navigator"
      Height          =   375
      Left            =   12240
      TabIndex        =   4
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate Thumbnails"
      Height          =   375
      Left            =   10560
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Remove Thumbnail"
      Height          =   375
      Left            =   12240
      TabIndex        =   3
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change"
      Height          =   375
      Left            =   13200
      TabIndex        =   2
      Top             =   6360
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   12240
      TabIndex        =   1
      Text            =   "200"
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "Split Annotation"
      Height          =   255
      Left            =   12480
      TabIndex        =   38
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Thumbnail Size"
      Height          =   255
      Left            =   12240
      TabIndex        =   37
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Thumbnails Per Row"
      Height          =   375
      Left            =   12480
      TabIndex        =   25
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LastThumb As CThumb
Dim nFrame As Long
Dim ca As New CAnnotations

Private Sub AnnotationCount_Click()
    If Not FreeImgEdit1.Annotations Is Nothing Then
        lblAnnotationCount.Caption = FreeImgEdit1.AnnotationCount
    End If
End Sub

Private Sub cmdGetCoordinates_Click()
    Dim a As CAnnotation
    
    If FreeImgEdit1.Selection.Count > 0 Then
        Set a = FreeImgEdit1.Selection(1)
        txtLeft.Text = a.Left
        txtTop.Text = a.Top
        txtWidth.Text = a.Width
        txtHeight.Text = a.Height
    End If
    
End Sub

Private Sub cmdGetThumb_HeightWidth_Click()
    txtThumbWidth.Text = FreeThumb1.ThumbnailWidth
    txtThumbHeight.Text = FreeThumb1.ThumbnailHeight
End Sub

Private Sub Command1_Click()
Dim sFolder As String
    sFolder = Text2.Text
    ReadFolder sFolder, Text3.Text
    FreeThumb1.Refresh
End Sub

Private Sub Command2_Click()
    FreeThumb1.MaxSize = Val(Text1.Text)
    FreeThumb1.Refresh
End Sub

Private Sub Command3_Click()
    FreeThumb1.Remove FreeThumb1.Selected
    FreeThumb1.Refresh
End Sub

Private Sub Command4_Click()
    FreeThumb1.ShowNavigator = Not FreeThumb1.ShowNavigator
End Sub

Private Sub Command5_Click()
    FreeThumb1.ThumbnailsPerRow = FreeThumb1.ThumbnailsPerRow - 1
    FreeThumb1.Refresh
End Sub

Private Sub Command6_Click()
    FreeThumb1.ThumbnailsPerRow = FreeThumb1.ThumbnailsPerRow + 1
    FreeThumb1.Refresh
End Sub

Private Sub Command7_Click()
    FreeImgEdit1.SplitAnnotation FreeImgEdit1.SelectedAnnotation, Split_Vertical
End Sub

Private Sub Command8_Click()
    ca.RemoveAt 1
  '  Dim a As CAnnotation
  '  For Each a In ca
  '      Debug.Print a.ToString
  '  Next
End Sub


Private Sub Command9_Click()
    FreeImgEdit1.SplitAnnotation FreeImgEdit1.SelectedAnnotation, Split_Horizontal
End Sub

Private Sub Form_Load()
    Me.Show
    'Set ca = FreeImgEdit1.Annotations
End Sub

Private Sub ReadFolder(path As String, Optional Ext As String = "*.tif")
Dim b()
Dim n
Dim T As CThumb

    a = Dir(path & "\" & Ext)
    While a <> ""
        ReDim Preserve b(n)
        b(n) = a
        a = Dir
        n = n + 1
    Wend
    
    For I = 0 To n - 1
        FreeThumb1.AddFile path & "\" & b(I)
        FreeThumb1.Refresh
        DoEvents
    Next
    
End Sub

Private Sub UpdateWindow()
    FreeImgEdit1.Display
    FreeThumb1.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set LastThumb = Nothing
End Sub

Private Sub FreeImgEdit1_AnnotationCreate(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, ByVal Rotation As Long, ByVal Page As Long)
    Dim thisAnnotation As New CAnnotation

    thisAnnotation.Left = Left
    thisAnnotation.Top = Top
    thisAnnotation.Width = Width
    thisAnnotation.Height = Height
    thisAnnotation.Page = Page
    thisAnnotation.Rotation = Rotation
    'thisAnnotation.LineWidth = 3
    'thisAnnotation.Filled = True
    'thisAnnotation.FillColor = RGB(Rnd() * 255, Rnd() * 255, Rnd() * 255)

    'thisAnnotation.Text = InputBox("Enter text here", "Hello world!")

    FreeImgEdit1.Annotations.Add thisAnnotation
End Sub

Private Sub FreeImgEdit1_BeforeAnnotationCreate(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, Cancel As Boolean, AnnotationType As FreeImageSuite.ANNOTATION_TYPE, FillColor As Long, LineColor As Long, Filled As Boolean)
    Debug.Print "Ping!"
End Sub

Private Sub FreeImgEdit1_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbCtrlMask) = vbCtrlMask Then
        Select Case KeyCode
        Case vbKeyAdd
            FreeImgEdit1.Zoom = FreeImgEdit1.Zoom * 1.5
            UpdateWindow
        Case vbKeySubtract
            FreeImgEdit1.Zoom = FreeImgEdit1.Zoom / 1.5
            UpdateWindow
        Case vbKeyPageUp
            FreeImgEdit1.RotateLeft
            UpdateWindow
        Case vbKeyPageDown
            FreeImgEdit1.RotateRight
            UpdateWindow
        End Select

    Else
        Select Case KeyCode
        Case vbKeyDelete
            If FreeImgEdit1.SelectedAnnotation > -1 Then
                FreeImgEdit1.RemoveAnnotation FreeImgEdit1.SelectedAnnotation
                FreeImgEdit1.Display
            End If
        
        Case vbKeyPageUp
            FreeImgEdit1.Page = FreeImgEdit1.Page - 1
            FreeImgEdit1.Display
        
        Case vbKeyPageDown
            FreeImgEdit1.Page = FreeImgEdit1.Page + 1
            FreeImgEdit1.Display
        
        Case vbKeyZ
            FreeImgEdit1.ZoomToAnnotation FreeImgEdit1.SelectedAnnotation
            FreeImgEdit1.Display
        End Select
    End If

End Sub

Private Sub FreeThumb1_ItemClick(Item As CThumb)
    If Not LastThumb Is FreeThumb1.Selected Then
        FreeImgEdit1.Image = Item.FileName

        Set FreeImgEdit1.Annotations = Item.Annotations
        FreeImgEdit1.SetNavigator Item
        
        'Set Item.Selection = FreeImgEdit1.Selection
        
        FreeImgEdit1.FitTo FIT_WIDTH

        FreeImgEdit1.Display
        Set LastThumb = FreeThumb1.Selected
    End If
End Sub

Private Sub FreeThumb1_ItemSelect(Item As CThumb)
    'Set ca = Item.Annotations
End Sub

Private Sub FreeThumb1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim AnnotationIndex As Long
Dim ThumbNailIndex As Long
    AnnotationIndex = -1
    ThumbNailIndex = -1
    If Shift = vbCtrlMask Then
        Select Case KeyCode
        Case vbKeyAdd
            FreeImgEdit1.Zoom = FreeImgEdit1.Zoom * 1.5
            FreeImgEdit1.Display
        Case vbKeySubtract
            FreeImgEdit1.Zoom = FreeImgEdit1.Zoom / 1.5
            FreeImgEdit1.Display
        Case vbKeyN
            FreeThumb1.ShowNavigator = Not FreeThumb1.ShowNavigator
            
        End Select
    Else
        Select Case KeyCode
        Case vbKeyDelete
        
            'ThumbNailIndex = FreeThumb1.Selected
            'AnnotationIndex = FreeThumb1.Thumbnails(ThumbNailIndex).SelectedAnnotation
            'If AnnotationIndex > -1 Then
            '    FreeThumb1.RemoveAnnotation ThumbNailIndex, AnnotationIndex
                'FreeImgEdit1.RemoveAnnotation FreeThumb1.Thumbnails(FreeThumb1.Selected).SelectedAnnotation
            '    FreeThumb1.Refresh
                'FreeImgEdit1.Display
            'End If
        End Select
    End If
End Sub

