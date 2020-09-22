VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private bBold As Boolean
Private bItalic As Boolean
Private bUnderlined As Boolean
Private hFont As Long
Private mFace As String
Private mSize As Long
Private mAngle As Integer
Private mDisposed As Boolean

Property Get Handle() As Long
    Handle = hFont
End Property

Property Let Face(v As String)
    mFace = v
End Property

Property Let Size(v As Long)
    mSize = v
End Property

Property Let Angle(v As Long)
    mAngle = v
End Property

Public Function CreateFont(Optional sFace As String, Optional lSize As Long, Optional iAngle As Long) As Long
Dim mLogFont As LOGFONT
Dim i As Long

    If Len(sFace) > 0 Then mFace = sFace
    If lSize > 0 Then mSize = lSize
    If Abs(iAngle) > 0 Then mAngle = iAngle
    
    Do While mAngle < 0
        mAngle = mAngle + 360
    Loop
    
    Do While mAngle >= 360
        mAngle = mAngle - 360
    Loop
    
    With mLogFont
        .lfEscapement = mAngle * 10
        .lfOrientation = mAngle * 10
        .lfItalic = bItalic
        .lfUnderline = bUnderlined
        
        If bBold Then
            .lfWeight = FW_BOLD
        Else
            .lfWeight = FW_NORMAL
        End If
        
        For i = 1 To Min(LF_FACESIZE - 1, Len(mFace))
            .lfFaceName(i) = Asc(Mid(mFace, i, 1))
        Next
        
        .lfHeight = mSize
    End With

    hFont = CreateFontIndirect(mLogFont)
    
    mDisposed = False
    
    CreateFont = hFont
End Function

Private Sub Class_Initialize()
    mFace = "Times New Roman"
    mSize = 14
End Sub

Public Sub Dispose()
    If hFont <> 0 Then
        DeleteObject hFont
        hFont = 0
    End If
    mDisposed = True
End Sub

Private Sub Class_Terminate()
    If Not mDisposed Then Dispose
End Sub
