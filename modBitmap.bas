Attribute VB_Name = "modBitmap"
'======================================================================
' Danny de Haas
' 29 April 2001
' Amsterdam, the Netherlands
' ddhadam@yahoo.com
'======================================================================

Option Explicit

' Declare some API Stuff
Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Const SRCCOPY = &HCC0020

Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function PlgBlt Lib "gdi32" (ByVal hdcDest As Long, lpPoint As POINTAPI, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hbmMask As Long, ByVal xMask As Long, ByVal yMask As Long) As Long

' Store Pictures in this format
Public Type tNEWBMP
    hdl As Long
    hdlOld As Long
    hdc As Long
    lWidth As Long
    lHeight As Long
End Type

Public Function FlipBitmap(ByRef hdcSource As Long, ByRef SourceWidth As Long, ByRef SourceHeight As Long) As tNEWBMP
    Dim pos(2) As POINTAPI
    Dim lrtnV As Long
    Dim lY As Long
    
    ' To flip the picture
    pos(0).X = 0
    pos(0).Y = SourceHeight - 1
    
    pos(1).X = SourceWidth
    pos(1).Y = SourceHeight - 1
    
    pos(2).X = 0
    pos(2).Y = -1
    
    ' Cleanup before beginning
    FreeResource FlipBitmap
    
    With FlipBitmap
        .hdc = CreateCompatibleDC(0)
        If .hdc <> 0 Then
            
            .hdl = CreateCompatibleBitmap(hdcSource, SourceWidth, SourceHeight)
            If .hdl <> 0 Then
                .hdlOld = SelectObject(.hdc, .hdl)

                lrtnV = PlgBlt(.hdc, pos(0), hdcSource, 0, 0, SourceWidth, SourceHeight, 0, 0, 0)
                If lrtnV = 0 Then
                    ' PlgBlt Failed to work
                    
                    ' Lets try to flip the picture in the old fashion way
                    For lY = 0 To SourceHeight
                    
                        lrtnV = BitBlt(.hdc, 0, lY, SourceWidth, 1, hdcSource, 0, (SourceHeight - lY - 1), SRCCOPY)
                        If lrtnV = 0 Then
                            ' It's failing again
                            
                            .hdc = 0
                            MsgBox "The bitmap is not build correctly.", vbCritical, "Error (PlgBlt & BitBlt)"
                            Exit For
                        End If
                        
                    Next lY
                End If
                
                .lWidth = SourceWidth
                .lHeight = SourceHeight
                
            Else
                MsgBox "The bitmap is not build correctly.", vbCritical, "Error (hdc)"
                
            End If
        Else
            MsgBox "The bitmap is not build correctly.", vbCritical, "Error (hdl)"
            
        End If
    End With
End Function

Public Sub FreeResource(Res As tNEWBMP)
    ' Free memory
    
    With Res
        SelectObject .hdc, .hdlOld
        DeleteObject .hdl
        DeleteObject .hdlOld
        DeleteDC .hdc
    End With
End Sub

