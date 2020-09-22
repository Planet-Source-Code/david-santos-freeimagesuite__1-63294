Attribute VB_Name = "mThumbThread"
Public sFilename As String

Public Sub ThreadFunc(Dummy As Variant)
Dim hTemp As Long
Dim hTemp2 As Long

    If hImage <> 0 Then
        FreeImage_Unload hImage
        mHasThumbnail = False
    End If
    
    If Dir(sFilename) <> "" Then
        
        T = GetTickCount
        hTemp = FreeImage_Load(FreeImage_GetFIFFromFilename(sFilename), sFilename)
        Debug.Print "Loading: " & Round((GetTickCount - T) / 1000, 2) & "s"
        
        If hTemp <> 0 Then
            T = GetTickCount
            
            mOrigWidth = FreeImage_GetWidth(hTemp)
            mOrigHeight = FreeImage_GetHeight(hTemp)
            
            Dim temp As New CMemoryDC
            temp.Height = mOrigHeight / 8
            temp.Width = mOrigWidth / 8

            FreeImage_PaintDCEx temp.hDC, hTemp, 0, 0, temp.Width, temp.Height, 0, 0, mOrigWidth, mOrigHeight
            FreeImage_Unload hTemp
            hTemp = FreeImage_CreateFromDC(temp.hDC)
            
            hImage = FreeImage_MakeThumbnail(hTemp, mSize)
            Set temp = Nothing
            
            Debug.Print "Creating: " & Round((GetTickCount - T) / 1000, 2) & "s"
            
            If hImage <> 0 Then
                mWidth = FreeImage_GetWidth(hImage)
                mHeight = FreeImage_GetHeight(hImage)
                
                vZoom = mWidth / mOrigWidth
                
                CreateThumbnail = True
                sFileTitle = Mid(sFilename, InStrRev(sFilename, "\") + 1)
                mHasThumbnail = True
            Else
                sError = "Error creating thumbnail"
            End If
        
            FreeImage_Unload hTemp
        Else
            sError = "Error loading image"
        End If
    Else
        sError = "File not found"
    End If
End Sub

