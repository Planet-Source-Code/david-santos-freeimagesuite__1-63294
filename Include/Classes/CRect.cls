VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Encapsulates a RECT that can be passed as a parameter and returned as an object
Private coords As RECT

Public Sub CopyRect(src As CRect)
    coords.Left = src.Left
    coords.Top = src.Top
    coords.Right = src.Right
    coords.Bottom = src.Bottom
End Sub

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
    Width = coords.Right - coords.Left
End Property

Property Get Height() As Long
    Height = coords.Bottom - coords.Top
End Property

Property Let Width(v As Long)
    coords.Right = coords.Left + v
End Property

Property Let Height(v As Long)
    coords.Bottom = coords.Top + v
End Property

Public Sub CreateRect(lTop As Long, lLeft As Long, lRight As Long, lBottom As Long)
    With coords
        .Top = lTop
        .Left = lLeft
        .Right = lRight
        .Bottom = lBottom
    End With
End Sub

Public Sub CreateRect2(lTop As Long, lLeft As Long, lWidth As Long, lHeight As Long)
    With coords
        .Top = lTop
        .Left = lLeft
        .Right = lLeft + lWidth
        .Bottom = lTop + lHeight
    End With
End Sub

