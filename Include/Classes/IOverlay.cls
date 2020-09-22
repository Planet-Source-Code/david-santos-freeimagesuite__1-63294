VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IOverlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Public Enum HitTestEnum
    HIT_NONE = 0
    HIT_LEFT = 1
    HIT_TOP = 2
    HIT_RIGHT = 4
    HIT_BOTTOM = 8
    HIT_CENTER = 16
    ' Special flags for lines (should not be mixed!)
    HIT_START = 32
    HIT_END = 64
    HIT_MOVE = 128
End Enum

Public Function GetRect() As CRect
End Function

Public Property Get AnnotationType() As ANNOTATION_TYPE
End Property

Public Property Let Selected(v As Boolean)
End Property

Public Property Get Selected() As Boolean
End Property

Public Sub Render(Container As IContainer)
End Sub

Public Sub Move(x, y, Container As IContainer)
End Sub

Public Sub Rotate(angle As Double)
End Sub

Public Function HitTest(x As Variant, y As Variant, Container As IContainer) As HitTestEnum
End Function

Property Let Top(v As Long)
End Property

Property Let Left(v As Long)
End Property

Property Let Bottom(v As Long)
End Property

Property Let Right(v As Long)
End Property

Property Let Width(v As Long)
End Property

Property Let Height(v As Long)
End Property

Property Let Rotation(v As Long)
End Property

Property Let Page(v As Long)
End Property

Property Get Top() As Long
End Property

Property Get Left() As Long
End Property

Property Get Bottom() As Long
End Property

Property Get Right() As Long
End Property

Property Get Width() As Long
End Property

Property Get Height() As Long
End Property

Property Get Rotation() As Long
End Property

Property Get Page() As Long
End Property

Property Let HitPoint(v As HitTestEnum)
End Property
