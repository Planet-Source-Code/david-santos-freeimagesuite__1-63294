VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileCopier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Event FileProgress(TotalFiles As Long, CurrentPosition As Long, filename As String, Percent As Long)
Public Event CopyFailed(filename As String)
Public Event Cancelled()
Public Event Complete()

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub Transfer(sDestination As String, Optional sSource As String = "", Optional sFilter As String = "")
    If Len(sSource) > 1 Then
        If Right(sSource, 1) = "\" Then sSource = Left(sSource, Len(sSource) - 1)
    End If
    
    If Right(sDestination, 1) = "\" Then sDestination = Left(sDestination, Len(sDestination) - 1)
    
    With frmCopier
        .Source = sSource
        .Destination = sDestination
        .Filter = sFilter
    End With
    
    frmCopier.Timer1.Enabled = True
End Sub

Private Sub Class_Initialize()
    Load frmCopier
    Set frmCopier.Parent = Me
End Sub

Property Let FileLimit(lCount As Long)
    If lCount < 1 Then lCount = 1
    frmCopier.lLimit = lCount
End Property

Property Get FileLimit() As Long
    FileLimit = frmCopier.lLimit
End Property

Property Let WaitTimeOut(lCount As Long)
    If lCount < 1 Then lCount = 1
    frmCopier.lLimit = lCount
End Property

Property Get WaitTimeOut() As Long
    FileLimit = frmCopier.lLimit
End Property

Public Sub Progress(TotalFiles As Long, CurrentPosition As Long, filename As String, Percent As Long)
    RaiseEvent FileProgress(TotalFiles, CurrentPosition, filename, Percent)
End Sub

Public Sub TransferComplete()
    RaiseEvent Complete
End Sub

Public Sub FileCopyFailed(filename As String)
    RaiseEvent CopyFailed(filename)
End Sub

Public Sub Continue()
    If FinishedCopying Then
        frmCopier.lTimeCtr = frmCopier.lTimeOut
        frmCopier.Timer2.Enabled = True
    End If
End Sub

Public Sub Cancel()
    bCancel = True
End Sub

Private Sub Class_Terminate()
    If bCancel Then RaiseEvent Cancelled
    
    Do While Not FinishedCopying
        DoEvents
    Loop
    
    Unload frmCopier
End Sub

