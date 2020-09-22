Attribute VB_Name = "basGeneral"
Option Explicit

Public Type POINTAPI
    X                              As Long
    Y                              As Long
End Type

Public Type BITMAP '14 bytes
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type



Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal X As Long, _
                                               ByVal Y As Long, _
                                               ByVal crColor As Long) As Long
Public Declare Function SleepEx Lib "Kernel32" (ByVal dwMilliseconds As Long, _
                                                 ByVal bAlertable As Long) As Long

Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal X As Long, _
                                               ByVal Y As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long

Public Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, _
                                                                            ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) As Long
Public Declare Function ReleaseCapture Lib "user32.dll" () As Long
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, _
                                                                            ByVal wMsg As Long, _
                                                                            ByVal wParam As Long, _
                                                                            lParam As Any) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, _
                                                                      ByVal crKey As Long, _
                                                                      ByVal bAlpha As Byte, _
                                                                      ByVal dwFlags As Long) As Long
                                                                      Public Declare Function timeGetTime Lib "winmm.dll" () As Long
                                                                      
Public Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, _
                                                           ByVal X As Long, _
                                                           ByVal Y As Long, _
                                                           ByVal destWidth As Long, _
                                                           ByVal destHeight As Long, _
                                                           ByVal hSrcDC As Long, _
                                                           ByVal xSrc As Long, _
                                                           ByVal ySrc As Long, _
                                                           ByVal nSrcWidth As Long, _
                                                           ByVal nSrcHeight As Long, _
                                                           ByVal crTransparent As Long) As Boolean
                                                           
Public Declare Function WindowFromPoint Lib "User32" (ByVal xPoint As Long, _
                                                       ByVal yPoint As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long

Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateDiscardableBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function GetDC Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long


Public blnBusy As Boolean

Public lngY(1000)               As Long
Public lngButtontoStyleIndex(100)                As Long
Public lngX(1000)               As Long

Dim blnBusy1 As Boolean

Public Sub LoadThemeFile(ByVal File As String, Skinimages() As StdPicture)
  ' This Function Describes How To Extract Data From .xbn File
  On Error GoTo Err
CheckBusy , , "LoadThemeFile"
CheckBusy , True, "LoadThemeFile"

blnBusy1 = True
  Dim nFileData As String
  Dim nPointer1 As Long
  Dim nData     As String
  Dim nPointer  As Long
  
    'pic_tmp1.Picture = Nothing
    ' Open File's One By One

    ReDim Skinimages(0)
'      Debug.Print "LoadFile"
'ReDim Captions(0)
    'If Dir(File) = "" Then Exit Sub':( --> replaced by:
    If LenB(Dir(File)) = 0 Then
        Exit Sub
    End If
    'Open Theme File
    nData = String$(FileLen(File), 0)
    Open File For Binary As #1
    Get #1, , nData
    Close #1
    'If Len(nData) < 100 Then Exit Sub':( --> replaced by:
    If Len(nData) < 100 Then
        Exit Sub
    End If
    'Remove The Header
    nData = Right$(nData, Len(nData) - 12)
    ' Reset Images
    'Search For Bmp Files
    nPointer1 = 1
re:
    For nPointer = nPointer1 To Len(nData) - 12
        ' Its Rubbish Header Data ( Unknown Functions ) If We Found This Header Above File Data Then
        If Asc(Mid$(nData, nPointer, 1)) > 0 Then
            If Asc(Mid$(nData, nPointer + 1, 1)) = 0 Then
                If Asc(Mid$(nData, nPointer + 2, 1)) = 0 Then
                    If Asc(Mid$(nData, nPointer + 3, 1)) = 0 Then
                        If Asc(Mid$(nData, nPointer + 4, 1)) >= 0 Then
                            If Asc(Mid$(nData, nPointer + 5, 1)) = 0 Then
                                If Asc(Mid$(nData, nPointer + 6, 1)) = 0 Then
                                    If Asc(Mid$(nData, nPointer + 7, 1)) = 0 Then
                                        If Asc(Mid$(nData, nPointer + 8, 1)) = 66 Then
                                            If Asc(Mid$(nData, nPointer + 9, 1)) = 77 Then
                                                'Found Header
                                                ' Get The Bmp File's Starting Point
                                                nFileData = Mid$(nData, nPointer, IIf(InStr(nPointer + 50, nData, "BM") > 0, InStr(nPointer + 10, nData, "BM") - 9, Len(nData)))
                                                ' Remove Some Extra Chars ( If Found ) Maybe Currupt Header
                                                nFileData = Mid$(nFileData, InStr(1, nFileData, "BM"))
                                                ' Save It
                                                Open "Button.tmp" For Binary As #1
                                                Put #1, , nFileData
                                                Close #1
                                                ' Load It
                                                'pic_tmp1.Picture =
                                                'Save it in Array
                                                ReDim Preserve Skinimages(UBound(Skinimages) + 1)
                                                lngX(UBound(Skinimages)) = Asc(Mid$(nData, nPointer + 4, 1))
                                                lngY(UBound(Skinimages)) = Asc(Mid$(nData, nPointer, 1))
                                                Set Skinimages(UBound(Skinimages)) = LoadPicture("Button.tmp")
                                                'pic_tmp1.Picture = Nothing
                                                'If Dir("Button.tmp") <> "" Then Kill "Button.tmp"':( --> replaced by:
                                                If Dir("Button.tmp") <> "" Then
                                                    Kill "Button.tmp"
                                                End If
                                                'Skip 10 Chars
                                                nPointer = nPointer + 50
                                                nPointer1 = InStr(nPointer, nData, "BM") - 15
                                                'If npointer1 > 0 Then GoTo re':( --> replaced by:
                                                If nPointer1 > 0 Then
                                                    GoTo re
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next nPointer

Err:
    Err.Clear
end1:
    nData = vbNullString
    blnBusy1 = False
'    Debug.Print "LoadFile End"
    With Err
      If .Number > 0 Then
            MsgBox "Error Loading this Theme File! " & vbCrLf & .Description
            .Clear
        End If
    End With 'Err
End Sub




Function GetMaskColor(objpicTemp As Long)
Dim MaskColor As Long
    'If GetPixel(objpicTemp.hDC, 0, 0) = GetPixel(objpicTemp.hDC, (objpicTemp.Width / 15) - 1, (objpicTemp.Height / 15) - 1) Then
    '    MaskColor = GetPixel(objpicTemp.hDC, (objpicTemp.Width / 15) - 1, (objpicTemp.Height / 15) - 1)
    'ElseIf GetPixel(objpicTemp.hDC, 0, 0) = GetPixel(objpicTemp.hDC, (objpicTemp.Width / 15) - 1, 0) Then 'NOT GETPIXEL(OBJpicTemp.HDC,...
    '    MaskColor = GetPixel(objpicTemp.hDC, (objpicTemp.Width / 15) - 1, 0)
    'ElseIf GetPixel(objpicTemp.hDC, 0, 0) = GetPixel(objpicTemp.hDC, 0, (objpicTemp.Height / 15) - 1) Then 'NOT GETPIXEL(OBJpicTemp.HDC,...
    '    MaskColor = GetPixel(objpicTemp.hDC, 0, (objpicTemp.Height / 15) - 1)
    'Else 'NOT GETPIXEL(OBJpicTemp.HDC,...
    '    MaskColor = RGB(255, 0, 255)
    'End If
      ' Transparent Mask Color
        MaskColor = GetPixel(objpicTemp, 0, 0)
    If MaskColor = RGB(255, 149, 255) Or MaskColor = RGB(152, 152, 152) Or MaskColor = RGB(148, 34, 158) Or MaskColor = RGB(177, 184, 194) Or MaskColor = RGB(255, 0, 0) Or MaskColor = RGB(0, 0, 0) Or MaskColor = RGB(255, 255, 255) Or MaskColor = RGB(255, 0, 255) Then
      Else 'NOT MASKCOLOR...
        MaskColor = RGB(255, 0, 255)
    End If

  
    GetMaskColor = MaskColor

End Function

Private Sub Resample(srcHDC As Long, _
                     OffsetX As Integer, _
                     OffsetY As Integer, _
                     ByVal srcW As Integer, _
                     ByVal srcH As Integer, _
                     ByVal dstW1 As Integer, _
                     ByVal dstH1 As Integer, _
                     dOffsetX As Integer, _
                     dOffsetY As Integer, _
                     DstHDC As Long, _
                     ByVal DstEdge As Byte)
Dim dx            As Integer
Dim dy            As Integer
Dim iX            As Integer
Dim iY            As Integer
Dim X             As Integer
Dim Y             As Integer
Dim i11           As Long
Dim i12           As Long
Dim i21           As Long
Dim i22           As Long
Dim V1            As Integer
Dim V2            As Integer
Dim V3            As Integer
Dim S1            As Integer
Dim S2            As Integer
Dim S3            As Integer
Dim N1            As Integer
Dim N2            As Integer
Dim N3            As Integer
Dim H1            As Integer
Dim H2            As Integer
Dim H3            As Integer
Dim U1            As Integer
Dim U2            As Integer
Dim U3            As Integer
Dim P1            As Integer
Dim P2            As Integer
Dim P3            As Integer
Dim Color11qRed   As Integer
Dim Color11qGreen As Integer
Dim Color11qBlue  As Integer
Dim Color21qRed   As Integer

Dim Color21qGreen As Integer
Dim Color21qBlue  As Integer
Dim Color22qRed   As Integer
Dim Color22qGreen As Integer
Dim Color22qBlue  As Integer
Dim Color12qRed   As Integer
Dim Color12qGreen As Integer
Dim Color12qBlue  As Integer
Dim dstW          As Integer
Dim dstH          As Integer
Dim iRX           As Integer
Dim iOrX          As Integer
Dim iRY           As Integer
Dim iOrY          As Integer
Dim dw            As Integer
Dim dh            As Integer
    blnBusy = True
    If srcHDC = 0 Then
        Exit Sub
    End If
    On Error Resume Next
    If DstEdge = 1 Then
        dstW = dstW1 + (dstW1 / srcW)
        dstH = dstH1 + (dstH1 / srcH)
    Else 'NOT DSTEDGE...
        dstW = dstW1
        dstH = dstH1
    End If
    For dy = 0 To srcH - 1
        iOrY = iRY
        iRY = ((dstH) / srcH) * (dy + 1)
        For dx = 0 To srcW - 1
            iOrX = iRX
            iRX = ((dstW) / srcW) * (dx + 1)
            '(Getting 4 Colors.  Of X, upper-left,
            'upper-right, lower-left, lower-right.)
            i11 = GetPixel(srcHDC, dx + OffsetX, dy + OffsetY)
            i12 = GetPixel(srcHDC, dx + 1 + OffsetX, dy + OffsetY)
            i21 = GetPixel(srcHDC, dx + OffsetX, dy + 1 + OffsetY)
            i22 = GetPixel(srcHDC, dx + 1 + OffsetX, dy + 1 + OffsetY)
            iX = iOrX
            iY = iOrY
            dw = iRX - iOrX
            dh = iRY - iOrY
            '(Get the Three Color values, Red, Green,
            'and blue.)
            '(upper-left)
            Color11qRed = i11 Mod 256
            Color11qGreen = (i11 \ 256) Mod 256
            Color11qBlue = (i11 \ 65536) Mod 256
            '(lower-left)
            Color12qRed = i12 Mod 256
            Color12qGreen = (i12 \ 256) Mod 256
            Color12qBlue = (i12 \ 65536) Mod 256
            '(upper-right)
            Color21qRed = i21 Mod 256
            Color21qGreen = (i21 \ 256) Mod 256
            Color21qBlue = (i21 \ 65536) Mod 256
            '(lower-right)
            Color22qRed = i22 Mod 256
            Color22qGreen = (i22 \ 256) Mod 256
            Color22qBlue = (i22 \ 65536) Mod 256
            '(Red)
            N1 = Color21qRed - Color11qRed
            H1 = Color11qRed
            '(Green)
            N2 = Color21qGreen - Color11qGreen
            H2 = Color11qGreen
            '(Blue)
            N3 = Color21qBlue - Color11qBlue
            H3 = Color11qBlue
            '(Cubic!)
            '(Red)
            U1 = Color22qRed - Color12qRed
            P1 = Color12qRed
            '(Green)
            U2 = Color22qGreen - Color12qGreen
            P2 = Color12qGreen
            '(Blue)
            U3 = Color22qBlue - Color12qBlue
            P3 = Color12qBlue
            For Y = 0 To dh - 1
                '(Now begins the Interpolation)
                Color11qRed = H1 + ((N1) / dh) * Y
                Color11qGreen = H2 + ((N2) / dh) * Y
                Color11qBlue = H3 + ((N3) / dh) * Y
                Color12qRed = P1 + ((U1) / dh) * Y
                Color12qGreen = P2 + ((U2) / dh) * Y
                Color12qBlue = P3 + ((U3) / dh) * Y
                '(Red)
                V1 = Color12qRed - Color11qRed
                S1 = Color11qRed
                '(Green)
                V2 = Color12qGreen - Color11qGreen
                S2 = Color11qGreen
                '(Blue)
                V3 = Color12qBlue - Color11qBlue
                S3 = Color11qBlue
                For X = 0 To dw - 1
                    Color11qRed = S1 + ((V1) / dw) * X
                    Color11qGreen = S2 + ((V2) / dw) * X
                    Color11qBlue = S3 + ((V3) / dw) * X
                    '(Set a Pixel, may need some changing,
                    If DstEdge = 1 Then
                        If X + iX < dstW1 Then
                            If Y + iY < dstH1 Then
                                SetPixel DstHDC, X + iX + dOffsetX, Y + iY + dOffsetY, RGB(Color11qRed, Color11qGreen, Color11qBlue)
                            End If
                        End If
                    Else 'NOT DSTEDGE...
                        SetPixel DstHDC, X + iX + dOffsetX, Y + iY + dOffsetY, RGB(Color11qRed, Color11qGreen, Color11qBlue)
                    End If
                Next X
            Next Y
            If dx = srcW - 1 Then
                iRX = 0
            End If
        Next dx
        If dy = srcH - 1 Then
            iRY = 0
        End If
    Next dy
    '''on error GoTo 0

End Sub

Public Sub ResampleBlt(destdc As Long, _
                       dx As Long, _
                       dy As Long, _
                       dw As Long, _
                       dh As Long, _
                       srcDC As Long, _
                       sx As Long, _
                       sy As Long, _
                       sw As Long, _
                       sh As Long, _
                       ByVal MaskColor As Long, Optional isResample As Boolean = True)
                       
If isResample = True Then
'If True = False Then
Resample srcDC, sx + 0, sy + 0, sw + 0, sh + 0, dw + 0, dh + 0, dx + 0, dy + 0, destdc, 0
'aTransparentBlt destdc, dx, dy, dw, dh, srcDC, sx, sy, sw, sh, 1, True
Else
TransparentBlt destdc, dx + 0, dy + 0, dw + 0, dh + 0, srcDC, sx + 0, sy + 0, sw + 0, sh + 0, MaskColor
'aTransparentBlt destdc, dx, dy, dw, dh, srcDC, sx, sy, sw, sh, 1, True
End If
End Sub



Public Sub RenderButton(picSrc As StdPicture, destHdc As Long, xSection As Long, ySection As Long, destWidth As Long, destHeight As Long, Optional isResample As Boolean = True, Optional BgColor As Long = &H8000000F)
Dim MaskColor As Long
Dim srcWidth  As Long
Dim srcHeight As Long
Dim srcHDC1 As Long
Dim hBmpOld1 As Long

Dim srcHDC As Long
Dim hBmp As Long
Dim hBmpOld As Long
Dim BMP As BITMAP

         If picSrc.Type = vbPicTypeBitmap Then
        hBmp = picSrc.Handle
       srcHDC = CreateCompatibleDC(0&)
        If srcHDC <> 0 Then
         hBmpOld = SelectObject(srcHDC, hBmp)
            If Not GetObject(hBmp, Len(BMP), BMP) <> 0 Then
             
             GoTo Err
             End If  'GetObject
       End If
    End If
   
   srcHDC1 = CreateCompatibleDC(0)
    hBmpOld1 = SelectObject(srcHDC1, CreateCompatibleBitmap(GetDC(0), destWidth / Screen.TwipsPerPixelX, destHeight / Screen.TwipsPerPixelY))  'CreateBitmap(destWidth / 15, destHeight / 15, 1, 1, 0))
    
'SetBkMode destHdc, 1
'SetBkColor destHdc, BgColor

SetBkMode srcHDC1, 1
SetBkColor srcHDC1, BgColor
FloodFill srcHDC1, 1, 1, BgColor


    
    'Convert Hmetric To Twips
    srcWidth = Int(((picSrc.Width * 567) / 1000))
    srcHeight = Int(((picSrc.Height * 567) / 1000))
    
    
    

MaskColor = GetMaskColor(srcHDC)
   

    
    'Set picDest.Picture = Nothing
    If destWidth <= 0 Then
    destWidth = 1000
    destHeight = 300
    End If
    
    If xSection <= 0 Then
        ySection = 60
        xSection = 60
    End If
    
   
xs:
    If (srcWidth - (xSection * 2)) / Screen.TwipsPerPixelX <= (srcWidth / 5) / Screen.TwipsPerPixelX Then
        xSection = xSection - 1
        GoTo xs
    End If
    xSection = Int(xSection)
    
ys:
    If (srcHeight - (ySection * 2)) / Screen.TwipsPerPixelX <= (srcHeight / 5) / Screen.TwipsPerPixelX Then
        ySection = ySection - 1
        GoTo ys
    End If
    ySection = Int(ySection)
    
    
    
      
            ResampleBlt srcHDC1, 0, 0, (xSection / Screen.TwipsPerPixelX), ySection / Screen.TwipsPerPixelY, srcHDC, 0, 0, xSection / Screen.TwipsPerPixelX, ySection / Screen.TwipsPerPixelY, MaskColor, isResample
            
            ResampleBlt srcHDC1, ((destWidth - xSection) / Screen.TwipsPerPixelX), 0, xSection / Screen.TwipsPerPixelX, ySection / Screen.TwipsPerPixelY, srcHDC, ((srcWidth - xSection) / Screen.TwipsPerPixelX), 0, xSection / Screen.TwipsPerPixelX, ySection / Screen.TwipsPerPixelY, MaskColor, isResample
            ResampleBlt srcHDC1, 0, ((destHeight - ySection)) / Screen.TwipsPerPixelX, xSection / Screen.TwipsPerPixelX, ySection / Screen.TwipsPerPixelY, srcHDC, 0, (srcHeight - ySection) / Screen.TwipsPerPixelX, xSection / Screen.TwipsPerPixelX, ySection / Screen.TwipsPerPixelY, MaskColor, isResample
            ResampleBlt srcHDC1, ((destWidth - xSection)) / Screen.TwipsPerPixelX, ((destHeight - ySection)) / Screen.TwipsPerPixelY, xSection / Screen.TwipsPerPixelX, ySection / Screen.TwipsPerPixelY, srcHDC, ((srcWidth - xSection) / Screen.TwipsPerPixelX), (srcHeight - ySection) / Screen.TwipsPerPixelY, (xSection / Screen.TwipsPerPixelX), ySection / Screen.TwipsPerPixelY, MaskColor, isResample
            'top right
            ResampleBlt srcHDC1, (xSection) / Screen.TwipsPerPixelX, 0, (((destWidth - (xSection * 2)) + 5) / Screen.TwipsPerPixelX), ySection / Screen.TwipsPerPixelY, srcHDC, xSection / Screen.TwipsPerPixelX, 0, ((srcWidth - (xSection * 2)) / Screen.TwipsPerPixelX), ySection / Screen.TwipsPerPixelY, MaskColor, isResample
            
            'Botom Center
            ResampleBlt srcHDC1, (xSection) / Screen.TwipsPerPixelX, ((destHeight - ySection)) / Screen.TwipsPerPixelY, (((destWidth - (xSection * 2)) + 5) / Screen.TwipsPerPixelX), ySection / Screen.TwipsPerPixelY, srcHDC, xSection / Screen.TwipsPerPixelX, (srcHeight - ySection) / Screen.TwipsPerPixelY, ((srcWidth - (xSection * 2)) / Screen.TwipsPerPixelX), ySection / Screen.TwipsPerPixelY, MaskColor, isResample
            
            'bottom right
            ResampleBlt srcHDC1, 0, (ySection) / Screen.TwipsPerPixelX, xSection / Screen.TwipsPerPixelX, ((destHeight - (ySection * 2)) + 5) / Screen.TwipsPerPixelX, srcHDC, 0, ySection / Screen.TwipsPerPixelY, xSection / Screen.TwipsPerPixelX, (srcHeight - (ySection * 2)) / Screen.TwipsPerPixelX, MaskColor, isResample
            
            'Right middle
            ResampleBlt srcHDC1, ((destWidth - xSection)) / Screen.TwipsPerPixelX, (ySection) / Screen.TwipsPerPixelY, xSection / Screen.TwipsPerPixelX, ((destHeight - (ySection * 2)) + 5) / Screen.TwipsPerPixelX, srcHDC, (((srcWidth) - xSection) / Screen.TwipsPerPixelX), ySection / Screen.TwipsPerPixelY, (xSection / Screen.TwipsPerPixelX), (srcHeight - (ySection * 2)) / Screen.TwipsPerPixelX, MaskColor, isResample
            
            
            ResampleBlt srcHDC1, (xSection) / Screen.TwipsPerPixelX, (ySection) / Screen.TwipsPerPixelX, ((destWidth - (xSection * 2)) + 5) / Screen.TwipsPerPixelY, ((destHeight - (ySection * 2)) + 5) / Screen.TwipsPerPixelX, srcHDC, xSection / Screen.TwipsPerPixelX, ySection / Screen.TwipsPerPixelY, ((srcWidth - (xSection * 2)) / Screen.TwipsPerPixelX), (srcHeight - (ySection * 2)) / Screen.TwipsPerPixelX, MaskColor, isResample
       

        'Another Pring Methord Not Used
           ' ResampleBlt srcHDC1, 0, 0, xSection / Screen.TwipsPerPixelX, (destHeight) / Screen.TwipsPerPixelY, srcHDC, 0, 0, xSection / Screen.TwipsPerPixelX, srcHeight / Screen.TwipsPerPixelY, MaskColor, isResample
           ' ResampleBlt srcHDC1, ((destWidth - xSection)) / Screen.TwipsPerPixelX, 0, (xSection / Screen.TwipsPerPixelX), destHeight / Screen.TwipsPerPixelY, srcHDC, (srcWidth - xSection) / Screen.TwipsPerPixelX, 0, xSection / Screen.TwipsPerPixelX, srcHeight / Screen.TwipsPerPixelX, MaskColor, isResample
           ' ResampleBlt srcHDC1, (xSection) / Screen.TwipsPerPixelX, 0, (((destWidth - (xSection * 2)) + 5) / Screen.TwipsPerPixelX), destHeight / Screen.TwipsPerPixelY, srcHDC, xSection / Screen.TwipsPerPixelX, 0, (srcWidth - (xSection * 2)) / Screen.TwipsPerPixelX, srcHeight / Screen.TwipsPerPixelX, MaskColor, isResample
        

    
        
TransparentBlt destHdc, 0, 0, destWidth / Screen.TwipsPerPixelX, destHeight / Screen.TwipsPerPixelY, srcHDC1, 0, 0, destWidth / Screen.TwipsPerPixelX, destHeight / Screen.TwipsPerPixelY, MaskColor

    
    
    
    
    
Call SelectObject(srcHDC1, hBmpOld1)
DeleteObject hBmpOld1
DeleteDC srcHDC1

Call SelectObject(srcHDC, hBmpOld)
DeleteObject hBmpOld
DeleteDC srcHDC


Err:
Err.Clear
End Sub
Public Sub CheckBusy(Optional lngTimetoWait As Long = 2, Optional Internal As Boolean = False, Optional SubName As String = "")
Dim A As String
A = IIf(blnBusy1 = True, Time & " General Busy " & SubName & " ", "") & "" & IIf(blnBusy = True, Time & " Sub Busy " & SubName, "")
If A <> "" Then Debug.Print A
Dim T As Long
T = timeGetTime + (lngTimetoWait * 1000)
If Internal = False Then
re:
DoEvents
If blnBusy = True And T > timeGetTime Then GoTo re
Else
re1:
DoEvents
If blnBusy1 = True And T > timeGetTime Then
GoTo re1

End If
End If
End Sub

