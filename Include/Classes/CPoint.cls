VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim mPoint As POINTAPI

Public Property Get x() As Long
    x = mPoint.x
End Property

Public Property Get y() As Long
    y = mPoint.y
End Property

Public Property Let x(v As Long)
    mPoint.x = v
End Property

Public Property Let y(v As Long)
    mPoint.y = v
End Property


