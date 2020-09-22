VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Counter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************************
'* Title   : All-Purpose Counter Class
'* Name    : Counter
'* Purpose : Provide an object that can be used with collections as named counters
'*           that can be dynamically allocated as needed.
'*           Such a counter can then be identified by its key and name, without the mess
'*           of declaring permanent variables.
'* Date    : Nov 30, 2005
'* Author  : Ender
'* Note    : May cause a slight performance hit
'**************************************************************************************

'**************************************************************************************
'* Added   : mCount2 variable and associated Property accessors
'* Date    : Nov 30, 2005
'* Purpose : Add an extra counter variable that can be linked to the same name
'* Note    : Should allow the user to set how many counters to be associated with a name
'**************************************************************************************
Option Explicit

Dim mName As String
Dim mCount As Long
Dim mCount2 As Long
Dim mTag As String

Property Let Name(sName As String)
    mName = sName
End Property

Property Get Name() As String
    Name = mName
End Property

Property Let Tag(sTag As String)
    mTag = sTag
End Property

Property Get Tag() As String
    Tag = mTag
End Property
 
Public Sub Increment(Optional amount As Long = 1)
    mCount = mCount + amount
End Sub

Public Sub Decrement(Optional amount As Long = 1)
    mCount = mCount - amount
End Sub

Public Sub Increment2(Optional amount As Long = 1)
    mCount2 = mCount2 + amount
End Sub

Public Sub Decrement2(Optional amount As Long = 1)
    mCount2 = mCount2 - amount
End Sub

Property Let Count(lCount As Long)
    mCount = lCount
End Property

Property Get Count() As Long
    Count = mCount
End Property

Property Let Count2(lCount As Long)
    mCount2 = lCount
End Property

Property Get Count2() As Long
    Count2 = mCount2
End Property


