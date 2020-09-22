VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFreeImageEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim mFileName As String
Dim mZoom As Single
Dim hImage As Long
Dim mBPP As Long
Dim mImageHeight As Long
Dim mImageWidth As Long
Dim mAutoRedraw As Boolean
Dim fmt As FREE_IMAGE_FORMAT
Dim lResX As Long
Dim lResY As Long
Dim lCurPage As Long
Dim lLastPage As Long
Dim lPageCount As Long
Dim hBitmap As Long
Dim sLastError As String

Property Get LastError() As String
    LastError = sLastError
End Property

Property Get Handle() As Long
    Handle = hImage
End Property

Property Get ImageWidth() As Long
    ImageWidth = mImageWidth
End Property

Property Get ImageHeight() As Long
    ImageHeight = mImageHeight
End Property

' *** Crop image to rectangle and Save as, without modifying the original image ***
Public Function CropAs(Left As Long, Top As Long, Right As Long, Bottom As Long, Filename As String, Optional Options As Long) As Boolean
Dim newImage As Long
    
    If mLeft < 0 Then mLeft = 0
    If mTop < 0 Then mTop = 0
    If mRight > mImageWidth Then mRight = mImageWidth
    If mBottom > mImageHeight Then mBottom = mImageHeight
    
    newImage = FreeImage_Copy(hImage, Left, Top, Right, Bottom)
    If newImage = 0 Then
        CropAs = False
        Exit Function
    End If
    
    If InStr(Filename, ".") = 0 Then
        ' do something?
    Else
        Select Case UCase(Mid(Filename, InStrRev(Filename, ".") + 1))
        Case "TIF", "TIFF"
            fmt = FIF_TIFF
        Case "BMP"
            fmt = FIF_BMP
        Case "GIF"
            fmt = FIF_GIF
        Case "JPG", "JPEG"
            fmt = FIF_JPEG
        Case "TGA"
            fmt = FIF_TARGA
        Case "PNG"
            fmt = FIF_PNG
        End Select
    End If
    
    FreeImage_SetDotsPerMeterX newImage, lResX
    FreeImage_SetDotsPerMeterY newImage, lResY
    
    'CropAs = FreeImage_Save(fmt, newImage, filename, Options) = CLng(True)
    Debug.Print FreeImage_Save(fmt, newImage, Filename, Options)
    CropAs = True
    FreeImage_Unload newImage
End Function

Public Function CropRotateAs(Left As Long, Top As Long, Right As Long, Bottom As Long, Rotation As Long, Filename As String, Options As Long) As Boolean
Dim newImage As Long
Dim newImage2 As Long
Dim tFmt As FREE_IMAGE_FORMAT
    
    newImage = FreeImage_Copy(hImage, Left, Top, Right, Bottom)

    If newImage <> 0 Then
        Select Case Rotation
        Case 90, -270
            newImage2 = FreeImage_RotateClassic(newImage, 90)
            FreeImage_Unload newImage
        Case -90, 270
            newImage2 = FreeImage_RotateClassic(newImage, -90)
            FreeImage_Unload newImage
        Case 180, -180
            newImage2 = FreeImage_RotateClassic(newImage, 180)
            FreeImage_Unload newImage
        Case 0, 360
            newImage2 = newImage
        End Select
    
        Select Case UCase(Mid(Filename, InStrRev(Filename, ".") + 1))
        Case "TIF", "TIFF"
            tFmt = FIF_TIFF
        Case "BMP"
            tFmt = FIF_BMP
        Case "GIF"
            tFmt = FIF_GIF
        Case "JPG", "JPEG"
            tFmt = FIF_JPEG
            Options = JPEG_QUALITYAVERAGE
        Case "TGA"
            tFmt = FIF_TARGA
        Case "PNG"
            tFmt = FIF_PNG
        End Select
        
        FreeImage_SetDotsPerMeterX newImage2, lResX
        FreeImage_SetDotsPerMeterY newImage2, lResY
        
        FreeImage_Save tFmt, newImage2, Filename, Options
        FreeImage_Unload newImage2
    
        CropRotateAs = True
        sLastError = ""
    Else
        sLastError = "Error cropping image " & mFileName & "."
        
        If (Left > mImageWidth) Or (Top > mImageHeight) Or (Left < 0) Or (Top < 0) Or (Right > mImageWidth) Or (Bottom > mImageHeight) Or (Right < 0) Or (Bottom < 0) Then
            sLastError = sLastError & " Coordinates out of bounds (" & Left & "," & Top & "," & Right & "," & Bottom & "). Image Size (" & mImageWidth & "x" & mImageHeight & ")"
        End If
        
        If (Right < Left) Or (Bottom < Top) Then
            sLastError = sLastError & " inverted or invalid coordinates (" & Left & "," & Top & "," & Right & "," & Bottom & ")"
        End If
        
        CropRotateAs = False
    End If
    
End Function

' *** Crop image to rectangle ***
Public Function Crop(Left As Long, Top As Long, Right As Long, Bottom As Long) As Boolean
Dim newImage As Long
    newImage = FreeImage_Copy(hImage, Left, Top, Right, Bottom)
    UpdateImage newImage
    Crop = newImage <> 0
End Function

' *** Image Rotation Functions ***
Public Function Flip() As Boolean
Dim newImage As Long
    newImage = FreeImage_RotateClassic(hImage, 180)
    UpdateImage newImage
    Flip = newImage <> 0
End Function

Public Function RotateLeft() As Boolean
Dim newImage As Long
    newImage = FreeImage_RotateClassic(hImage, 90)
    UpdateImage newImage
    RotateLeft = newImage <> 0
End Function

Public Function RotateRight() As Boolean
Dim newImage As Long
    newImage = FreeImage_RotateClassic(hImage, -90)
    UpdateImage newImage
    RotateRight = newImage <> 0
End Function

Public Sub Invert()
Dim newImage As Long
    FreeImage_Invert hImage
End Sub

Property Let Image(Filename)
On Error GoTo Image_Error
Dim tempImage As Long

    If hBitmap <> 0 Then
        FreeImage_CloseMultiBitmap hBitmap
        hBitmap = 0
    End If
    
    If hImage <> 0 Then
        FreeImage_Unload hImage
        hImage = 0
    End If
    
    lCurPage = 0
    lPageCount = 0
    
    ' blank assumes we want to unload the image from memory
    If Filename = "" Then Exit Property
    
    mFileName = Filename
    
    ' Set format based on extension
    Select Case UCase(Mid(mFileName, InStrRev(mFileName, ".") + 1))
    Case "TIF", "TIFF"
        fmt = FIF_TIFF
    Case "BMP"
        fmt = FIF_BMP
    Case "GIF"
        fmt = FIF_GIF
    Case "JPG", "JPEG"
        fmt = FIF_JPEG
    Case "TGA"
        fmt = FIF_TARGA
    Case "PNG"
        fmt = FIF_PNG
    End Select
    
    ' check for multi-page formats
    If (fmt = FIF_TIFF) Or (fmt = FIF_GIF) Then
                    
        ' try to load the file
        hBitmap = FreeImage_OpenMultiBitmap(fmt, Filename, CLng(False), CLng(False))
        
        If hBitmap <> 0 Then
            lPageCount = FreeImage_GetPageCount(hBitmap)
            ' get the first page
            lCurPage = 0
            tempImage = FreeImage_LockPage(hBitmap, lCurPage)
            ' make a copy of this page so we can rotate, etc...
            hImage = FreeImage_Clone(tempImage)
            ' return the page to the bitmap
            FreeImage_UnlockPage hBitmap, tempImage, CLng(False)
        End If
    
    Else
        ' single page image
        lPageCount = 1
        hImage = FreeImage_Load(fmt, mFileName)
    
    End If
    
    '_Load or _OpenMultiBitmap will set fmt to FIF_UNKNOWN
    Select Case fmt
    Case FIF_UNKNOWN
        Err.Raise vbObjectError + 6, "FreeImageEdit.Image (Property)", "FreeImageEdit.Image (Property): Sorry, I don't recognize this file."
        
    Case Else
        
        If hImage <> 0 Then
            mBPP = FreeImage_GetBPP(hImage)
            
            mImageWidth = FreeImage_GetWidth(hImage)
            mImageHeight = FreeImage_GetHeight(hImage)
            
            lResX = FreeImage_GetDotsPerMeterX(hImage)
            lResY = FreeImage_GetDotsPerMeterY(hImage)
            
            mScrollSizeX = mImageWidth * 0.05
            mScrollSizeY = mImageHeight * 0.05

        Else
            Err.Raise vbObjectError + 6, "FreeImageEdit.Image (Property)", "Error opening image"
        
        End If
    
    End Select
    
    ' reset rotation
    mRotation = 0

    On Error GoTo 0
    Exit Property

Image_Error:
    Err.Raise Err.Number, "Property Image Let in FreeImgEdit", Err.Description

End Property

Property Get Image()
    Image = mFileName
End Property

Public Function CloneFrom(newImgEdit As CFreeImageEdit) As Boolean
Dim newImage As Long

    newImage = FreeImage_Clone(newImgEdit.Handle)
    UpdateImage newImage
    Clone = newImage <> 0

    lResX = FreeImage_GetDotsPerMeterX(newImgEdit.Handle)
    lResY = FreeImage_GetDotsPerMeterY(newImgEdit.Handle)
    
End Function

Public Function Save() As Boolean
    Save = FreeImage_Save(fmt, hImage, mFileName) = 0
End Function

Public Function SaveAs(Filename As String, Optional Options As Long = 0) As Boolean
    If InStr(Filename, ".") = 0 Then
        ' do something?
    Else
        Select Case UCase(Mid(Filename, InStrRev(Filename, ".") + 1))
        Case "TIF", "TIFF"
            fmt = FIF_TIFF
        Case "BMP"
            fmt = FIF_BMP
        Case "GIF"
            fmt = FIF_GIF
        Case "JPG", "JPEG"
            fmt = FIF_JPEG
            Options = JPEG_QUALITYAVERAGE
        Case "TGA"
            fmt = FIF_TARGA
        Case "PNG"
            fmt = FIF_PNG
        End Select
    End If
    
    FreeImage_SetDotsPerMeterX hImage, lResX
    FreeImage_SetDotsPerMeterY hImage, lResY
    
    SaveAs = FreeImage_Save(fmt, hImage, Filename, Options) = 0
    If SaveAs Then mFileName = Filename
End Function

Property Let ResolutionX(Value As Long)
    lResX = Value
    FreeImage_SetDotsPerMeterX hImage, lResX
End Property

Property Let ResolutionY(Value As Long)
    lResY = Value
    FreeImage_SetDotsPerMeterY hImage, lResY
End Property

Property Get ResolutionX() As Long
    ResolutionX = lResX
End Property

Property Get ResolutionY() As Long
    ResolutionY = lResY
End Property

' *** Updates the in-memory image after image-destroying operations such as rotate
Private Sub UpdateImage(hHandle As Long)
    If hHandle <> 0 Then
        If hImage <> 0 Then FreeImage_Unload hImage
        
        ' point hImage to new image
        hImage = hHandle
        
        ' update size variables
        mBPP = FreeImage_GetBPP(hImage)
        mImageWidth = FreeImage_GetWidth(hImage)
        mImageHeight = FreeImage_GetHeight(hImage)
    
    End If
End Sub

Property Get Page() As Long
    Page = lCurPage
End Property

Property Let Page(lPage As Long)
Dim newImage As Long
Dim tempImage As Long
Dim lCount As Long
    
    lCurPage = lPage - 1
    
    If lCurPage > lPageCount - 1 Then lCurPage = lPageCount - 1
    If lCurPage < 0 Then lCurPage = 0
    
    If hBitmap <> 0 Then
        tempImage = FreeImage_LockPage(hBitmap, lCurPage)
        newImage = FreeImage_Clone(tempImage)
        FreeImage_UnlockPage hBitmap, tempImage, CLng(False)
        UpdateImage newImage
    End If
End Property

Property Get PageCount() As Long
    PageCount = lPageCount
End Property

Private Sub Class_Terminate()
    If hBitmap <> 0 Then
        FreeImage_CloseMultiBitmap hBitmap
        hBitmap = 0
    End If

    If hImage <> 0 Then
        FreeImage_Unload hImage
        hImage = 0
    End If
End Sub
