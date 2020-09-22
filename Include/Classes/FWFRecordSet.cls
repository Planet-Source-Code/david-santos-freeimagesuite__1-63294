VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FWFRecordSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************************
'* Name    : FWFRecordset
'* Title   : Fixed-Width Field Recordset Class
'* Purpose : Provide an easy-to-use interface for reading and writing records and mFields
'*           to and from a fixed-width field data source
'* Author  : Ender
'* Date    : Nov 28, 2005
'**************************************************************************************

'**************************************************************************************
'* ToDo    : Make a hash function for accessing mFields by name
'*         :
'**************************************************************************************

'**************************************************************************************
'* Added   : Counters
'* Date    : Nov 30, 2005
'* Purpose : This is to be able to monitor how many characters were read or written
'* Note    : Requires the Counter Class
'*         : Should allow the user to enable/disable counters, to improve speed
'**************************************************************************************
#Const ENABLE_COUNTERS = False

Option Explicit

Private Type FIELD_DEF
    Name As String
    Size As Long
End Type

Private Enum RECORD_MODE
    MODE_READ = 0
    MODE_ADD = 1
    MODE_EDIT = 2
End Enum

Private Type RECORD_DEF
    Field() As String
End Type

Dim mFields() As FIELD_DEF
Dim lFieldCount As Long

Dim Records() As RECORD_DEF
Dim lRecordCount As Long
Dim lRecordPointer As Long
Dim lRecordWidth As Long
Dim mFilename As String
Dim RecMode As RECORD_MODE

Private Const sInternalName = "FWFRecordset"

#If ENABLE_COUNTERS Then
Dim mCounters As New Collection
#End If

Public Sub AddField(Name As String, Size As Long)
    ReDim Preserve mFields(lFieldCount)
    mFields(lFieldCount).Name = Name
    mFields(lFieldCount).Size = Size
    lRecordWidth = lRecordWidth + Size
    
#If ENABLE_COUNTERS Then
    AddCounter Name, "Field" & lFieldCount
#End If

    lFieldCount = lFieldCount + 1
End Sub

Public Sub Clear()
    lRecordCount = 0
    Erase Records
    lRecordPointer = -1
End Sub

Public Sub AddNew()
    If RecMode = MODE_READ Then
        RecMode = MODE_ADD
        lRecordCount = lRecordCount + 1
        ReDim Preserve Records(lRecordCount - 1)
        lRecordPointer = lRecordCount - 1
        ReDim Records(lRecordPointer).Field(lFieldCount - 1)
    Else
        Err.Raise vbObjectError + 1, "AddNew", "Record is currently open"
    End If

End Sub

Public Sub Edit()
    If RecMode = MODE_READ Then
        RecMode = MODE_EDIT
        ' save current record ...
    Else
        Err.Raise vbObjectError + 1, "Edit", "Record is currently open"
    End If
End Sub

Public Sub Update()
    Select Case RecMode
    Case MODE_READ
        Err.Raise vbObjectError + 1, "CancelUpdate", "Record is not currently open"

    Case MODE_ADD
        RecMode = MODE_READ

    Case MODE_EDIT
        RecMode = MODE_READ

    End Select
End Sub

Public Sub CancelUpdate()
    Select Case RecMode
    Case MODE_READ
        Err.Raise vbObjectError + 1, "CancelUpdate", "Record is not currently open"

    Case MODE_ADD
        lRecordCount = lRecordCount - 1
        If lRecordCount > 0 Then
        ReDim Preserve Records(lRecordCount - 1)
        End If
        RecMode = MODE_READ

    Case MODE_EDIT
        RecMode = MODE_READ

        ' ... restore saved record

    End Select

End Sub

Public Sub Load(filename As String, Optional dWidth As Long = 0)
Dim DTALines() As String
Dim sBuffer As String
Dim hFile As Long
Dim ptr As Long
Dim i As Long, j As Long
Dim sFileTitle As String

    sFileTitle = Mid(filename, InStrRev(filename, "\") + 1)

'    If lDataWidth <= 0 Then
'        Err.Raise vbObjectError + 8, "FWFRecordset.Load", "Invalid record width specified: " & lDataWidth
'        Exit Sub
'    End If
    
    hFile = FreeFile
    
    lRecordCount = 0
    
    mFilename = Mid$(filename, InStrRev(filename, "\") + 1)
    
    
    Open filename For Input As hFile
    Do While Not VBA.EOF(hFile)
        Line Input #hFile, sBuffer
        If Len(sBuffer) > 0 Then
            
            If dWidth > 0 Then
                If Len(sBuffer) <> dWidth Then
                    'Err.Raise vbObjectError + 8, "FWFRecordset.Load", "Error while loading " & filename & " on Line " & (lRecordCount + 1) & "." & vbCrLf & vbCrLf & " Incorrect record width=" & Len(sBuffer) & ". Expected=" & dWidth
                    MsgBox "Error while loading " & sFileTitle & " on Line " & (lRecordCount + 1) & "." & vbCrLf & vbCrLf & " Incorrect record width=" & Len(sBuffer) & ". Expected=" & dWidth + 1, vbExclamation, "FWFRecordset.Load"
                    Exit Do
                End If
            End If
            
            If InStr(1, sBuffer, vbTab) > 0 Then
                MsgBox "TAB character was found in " & sFileTitle & " on Line " & (lRecordCount + 1) & "." & vbCrLf & "This may cause data errors", vbExclamation, "FWFRecordset.Load"
            End If
            
            ReDim Preserve DTALines(lRecordCount)
            DTALines(lRecordCount) = sBuffer
            lRecordCount = lRecordCount + 1
        End If
    Loop
    Close hFile
    
    If lRecordCount = 0 Then Exit Sub
    
    ReDim Records(lRecordCount - 1)
    
    'read each line
    For i = 0 To UBound(DTALines)
        ' read each field
        ReDim Records(i).Field(lFieldCount - 1)
        ptr = 1
        
        For j = 0 To UBound(mFields)
            Records(i).Field(j) = RTrim$(Mid$(DTALines(i), ptr, mFields(j).Size))
#If ENABLE_COUNTERS Then
            mCounters("Field" & j).Increment Len(Records(i).Field(j))
#End If
            ptr = ptr + mFields(j).Size
        Next
    Next
    
    lRecordPointer = 0
End Sub

Public Sub Save(filename As String)
    Dim sBuffer As String
    Dim i As Long, j As Long
    Dim ptr As Long

    ' added sanity check to make sure we actually have something to output
    If lRecordCount > 0 Then

        sBuffer = Space(lRecordCount * (lRecordWidth + 2))
    
        ptr = 1
        For i = 0 To UBound(Records)
            For j = 0 To UBound(mFields)
                Mid$(sBuffer, ptr, mFields(j).Size) = Left(Records(i).Field(j), mFields(j).Size)
                ptr = ptr + mFields(j).Size
            Next
    
            Mid$(sBuffer, ptr, 2) = vbCrLf
            ptr = ptr + 2
        Next
    
        ' ugly kludge to make sure we don't have an extra CrLf at the end
        sBuffer = Left$(sBuffer, Len(sBuffer) - 2)

    End If

    Open filename For Output As 1
    Print #1, sBuffer
    Close 1

    sBuffer = ""

End Sub

Property Get RecordCount() As Long
    RecordCount = lRecordCount
End Property

Property Get Field(index) As String
Attribute Field.VB_UserMemId = 0
Dim idx As Long
Dim i As Long

    If IsNumeric(index) Then
        If index < 1 Or index > lFieldCount Then
            Err.Raise vbObjectError + 1, sInternalName, "Invalid index (""" & index & """)"
            Exit Property
        End If
        Field = Records(lRecordPointer).Field(index - 1)
    Else
        idx = -1
        For i = 0 To UBound(mFields)
            If UCase(mFields(i).Name) = UCase(index) Then
                idx = i
                Exit For
            End If
        Next
        If idx > -1 Then
            Field = Records(lRecordPointer).Field(idx)
        Else
            Err.Raise vbObjectError + 1, sInternalName, "Field not foud (""" & index & """)"
        End If
    End If
End Property

Property Let Field(index, value As String)
Dim idx As Long
Dim i As Long

    If IsNumeric(index) Then
        If index < 1 Or index > lFieldCount Then
            Err.Raise vbObjectError + 1, sInternalName, "Invalid index (""" & index & """)"
            Exit Property
        End If
        Records(lRecordPointer).Field(index - 1) = value
    Else
        idx = -1
        For i = 0 To UBound(mFields)
            If UCase(mFields(i).Name) = UCase(index) Then
                idx = i
                Exit For
            End If
        Next
        If idx > -1 Then
            Records(lRecordPointer).Field(idx) = value
        Else
            Err.Raise vbObjectError + 1, sInternalName, "Field not foud (""" & index & """)"
        End If
    End If
End Property

Public Sub Move(pos)
Dim newpos As Long
    newpos = lRecordPointer + pos
    If newpos < 1 Or newpos > lRecordCount Then
        Err.Raise vbObjectError + 3, sInternalName, "Invalid position"
    End If
    lRecordPointer = newpos
End Sub

Property Get Fields(index) As String
    Fields = mFields(index).Name
End Property

Property Get FieldCount() As Long
    FieldCount = UBound(mFields)
End Property

Public Sub Sort(ParamArray varmFields() As Variant)
Dim i As Long, j As Long

    For i = UBound(varmFields) To 0 Step -1
        For j = 0 To UBound(mFields)
            If mFields(j).Name = varmFields(i) Then
                SortField j
                Exit For
            End If
        Next
    Next
End Sub

Private Sub SortField(index As Long, Optional descending As Boolean = False)
    Dim value As Variant, temp As RECORD_DEF
    Dim sp As Integer
    Dim leftStk(32) As Long, rightStk(32) As Long
    Dim leftNdx As Long, rightNdx As Long
    Dim i As Long, j As Long
    
    Dim numEls As Long
    
    ' account for optional arguments
    numEls = UBound(Records)
    ' init pointers
    leftNdx = LBound(Records)
    rightNdx = numEls
    ' init stack
    sp = 1
    leftStk(sp) = leftNdx
    rightStk(sp) = rightNdx

    Do
        If rightNdx > leftNdx Then
            value = Records(rightNdx).Field(index)
            i = leftNdx - 1
            j = rightNdx
            ' find the pivot item
            If descending Then
                Do
                    Do: i = i + 1: Loop Until Records(i).Field(index) <= value
                    Do: j = j - 1: Loop Until j = leftNdx Or Records(j).Field(index) >= value
                    temp = Records(i)
                    Records(i) = Records(j)
                    Records(j) = temp
                Loop Until j <= i
            Else
                Do
                    Do: i = i + 1: Loop Until Records(i).Field(index) >= value
                    Do: j = j - 1: Loop Until j = leftNdx Or Records(j).Field(index) <= value
                    temp = Records(i)
                    Records(i) = Records(j)
                    Records(j) = temp
                Loop Until j <= i
            End If
            ' swap found items
            temp = Records(j)
            Records(j) = Records(i)
            Records(i) = Records(rightNdx)
            Records(rightNdx) = temp
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

Public Sub Printout()
Dim sRec As String
Dim sField As String
Dim i As Long
Dim j As Long

    For i = 0 To UBound(Records)
        sRec = ""
        For j = 0 To UBound(mFields)
            sField = Space(mFields(j).Size)
            Mid(sField, 1) = Records(i).Field(j)
            sRec = sRec & sField
        Next
        Debug.Print sRec
    Next
    
End Sub

Public Sub MoveFirst()
    lRecordPointer = 0
End Sub

Public Sub MovePrev()
    lRecordPointer = lRecordPointer - 1
End Sub

Public Sub MoveNext()
    lRecordPointer = lRecordPointer + 1
End Sub

Public Sub MoveLast()
    lRecordPointer = lRecordCount - 1
End Sub

Property Get BOF() As Boolean
    BOF = (lRecordPointer < 0)
End Property

Property Get EOF() As Boolean
    EOF = (lRecordPointer = lRecordCount)
End Property

Property Get AbsoluteRecordPosition() As Long
    AbsoluteRecordPosition = lRecordPointer + 1
End Property

'**************************************************************************************
'* Counters Methods & Properties
'**************************************************************************************
#If ENABLE_COUNTERS Then

Public Sub AddCounter(Name, key)
Dim newCounter As New Counter
    newCounter.Name = Name
    mCounters.Add newCounter, key
End Sub

Public Sub ClearCounters()
Dim rCounter As Counter
    For Each rCounter In mCounters
        rCounter.Count = 0
    Next
End Sub

Property Get Counters()
    Set Counters = mCounters
End Property

#End If
