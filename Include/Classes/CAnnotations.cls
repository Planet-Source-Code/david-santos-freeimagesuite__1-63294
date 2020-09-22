VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAnnotations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"CAnnotation"
Attribute VB_Ext_KEY = "Member0" ,"CAnnotation"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable to hold collection
Private mCol As Collection
Private bUpdates As Boolean
Public Event Updated()
Private mName As String

Property Get Name() As String
    Name = mName
End Property

Public Function Clear()
    Do While mCol.Count > 0
        mCol.Remove 1
    Loop
    If bUpdates Then RaiseEvent Updated
End Function

Public Function Add(Annotation As Object, Optional Key As String) As Object
    'create a new object
    'Dim objNewMember As CAnnotation
    'Set objNewMember = New CAnnotation


    'set the properties passed into the method
    If Len(sKey) = 0 Then
        mCol.Add Annotation
    Else
        mCol.Add Annotation, Key
    End If


    'return the object created
    Set Add = Annotation
    'Set objNewMember = Nothing

    If bUpdates Then RaiseEvent Updated
End Function

Public Property Get Item(vntIndexKey As Variant) As Object
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property

Public Sub Remove(target As Object)
Dim ctr As Long
Dim obj As Object
    ctr = 1

    For Each obj In mCol
        If obj Is target Then
            mCol.Remove ctr
            Exit For
        End If
        ctr = ctr + 1
    Next
    
    RaiseEvent Updated
End Sub

Public Sub RemoveAt(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey

    If bUpdates Then RaiseEvent Updated
End Sub

Friend Sub CallUpdate()
    RaiseEvent Updated
End Sub

Property Get Updates() As Boolean
    Updates = bUpdates
End Property

Property Let Updates(v As Boolean)
    bUpdates = v
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
     Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
    bUpdates = True
    MyCounter = MyCounter + 1
    mName = "Collection" & MyCounter
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Debug.Print "Destroying " & mName
    Set mCol = Nothing
End Sub

