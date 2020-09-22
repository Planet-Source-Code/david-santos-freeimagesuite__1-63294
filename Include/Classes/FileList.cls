VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************************
'* Title   : FileList Class
'* Name    : FileList
'* Author  : Ender
'* Date    : Nov 28, 2005
'* Purpose : Provide a alternative to the FileList control that can be used without the
'*           need for a form
'**************************************************************************************

'**************************************************************************************
'* Added   : QuickSort procedure
'* Purpose : Sorts the file list
'**************************************************************************************

Option Explicit
Dim mSorted As Boolean
Dim mPath As String
Dim mFilter
Dim mAutoScan As Boolean
Dim asFileList() As String
Dim lFileListCount As Long
Dim lFileListMax As Long

Private Const sInternalName = "FileList Class"

' From http://www.devx.com/vb2themax/Tip/18973
Private Sub QuickSort(arr As Variant, Optional numEls As Variant, _
                      Optional descending As Boolean)

    Dim Value As Variant, temp As Variant
    Dim sp As Integer
    Dim leftStk(32) As Long, rightStk(32) As Long
    Dim leftNdx As Long, rightNdx As Long
    Dim i As Long, j As Long

    ' account for optional arguments
    If IsMissing(numEls) Then numEls = UBound(arr)
    ' init pointers
    leftNdx = LBound(arr)
    rightNdx = numEls
    ' init stack
    sp = 1
    leftStk(sp) = leftNdx
    rightStk(sp) = rightNdx

    Do
        If rightNdx > leftNdx Then
            Value = arr(rightNdx)
            i = leftNdx - 1
            j = rightNdx
            ' find the pivot item
            If descending Then
                Do
                    Do: i = i + 1: Loop Until arr(i) <= Value
                    Do: j = j - 1: Loop Until j = leftNdx Or arr(j) >= Value
                    temp = arr(i)
                    arr(i) = arr(j)
                    arr(j) = temp
                Loop Until j <= i
            Else
                Do
                    Do: i = i + 1: Loop Until arr(i) >= Value
                    Do: j = j - 1: Loop Until j = leftNdx Or arr(j) <= Value
                    temp = arr(i)
                    arr(i) = arr(j)
                    arr(j) = temp
                Loop Until j <= i
            End If
            ' swap found items
            temp = arr(j)
            arr(j) = arr(i)
            arr(i) = arr(rightNdx)
            arr(rightNdx) = temp
            ' push on the stack the pair of pointers that differ most
            sp = sp + 1
            If (i - leftNdx) > (rightNdx - i) Then
                leftStk(sp) = leftNdx
                rightStk(sp) = i - 1
                leftNdx = i + 1
            Else
                leftStk(sp) = i + 1
                rightStk(sp) = rightNdx
                rightNdx = i - 1
            End If
        Else
            ' pop a new pair of pointers off the stacks
            leftNdx = leftStk(sp)
            rightNdx = rightStk(sp)
            sp = sp - 1
            If sp = 0 Then Exit Do
        End If
    Loop
End Sub

Property Let Sorted(bSort As Boolean)
    mSorted = bSort
    If mSorted And lFileListCount > 0 Then QuickSort asFileList
End Property

Property Get Sorted() As Boolean
    Sorted = mSorted
End Property

Property Get FileCount() As Long
    FileCount = lFileListCount
End Property

Property Let Filter(sFilter As String)
    mFilter = sFilter
    If mAutoScan Then
        Clear
        ScanPath
    End If
End Property

Property Get Filter() As String
    Filter = mFilter
End Property

Public Sub Scan()
    Clear
    ScanPath
End Sub

Private Sub ScanPath()
    Dim a As String
    On Error GoTo ScanPath_Error

    If Len(Dir(mPath, vbDirectory)) > 0 Then
        a = Dir(mPath & mFilter)
        Do While a <> ""
            Add a
            a = Dir
        Loop
        If lFileListCount > 0 Then
            ReDim Preserve asFileList(lFileListCount - 1)
            If mSorted And lFileListCount > 0 Then QuickSort asFileList
        End If
    Else
        Err.Raise vbObjectError + 2, sInternalName, "Path not found"
    End If

    On Error GoTo 0
    Exit Sub

ScanPath_Error:
    Err.Raise vbObjectError + 6, "FileList", "Error scanning path. Path may be inaccessible"
End Sub

Property Let Path(sPath As String)
    Dim a As String
    mPath = sPath
    If Right(mPath, 1) <> "\" Then mPath = mPath & "\"
    If mAutoScan Then
        Clear
        ScanPath
    End If
End Property

Property Get Path() As String
    Path = mPath
End Property

Public Sub Add(filename As String)
    asFileList(lFileListCount) = filename
    lFileListCount = lFileListCount + 1
    If lFileListCount > lFileListMax Then
        lFileListMax = lFileListMax * 2
        ReDim Preserve asFileList(lFileListMax)
    End If
End Sub

Property Get File(Index) As String
    If Index > lFileListCount - 1 Or Index < 0 Then
        Err.Raise vbObjectError + 1, sInternalName, "Invalid index"
        Exit Property
    End If
    File = asFileList(Index)
End Property

Property Let File(Index, filename As String)
    If Index > lFileListCount - 1 Or Index < 0 Then
        Err.Raise vbObjectError + 1, sInternalName, "Invalid index"
        Exit Property
    End If
    asFileList(Index) = filename
End Property

Public Sub Remove(Index)
    Dim i As Long
    For i = Index To lFileListCount - 2
        asFileList(i) = asFileList(i + 1)
    Next
End Sub

Public Sub Clear()
    Erase asFileList
    lFileListMax = 500
    ReDim asFileList(lFileListMax)
    lFileListCount = 0
End Sub

Private Sub Class_Initialize()
    Clear
    mSorted = True
End Sub
