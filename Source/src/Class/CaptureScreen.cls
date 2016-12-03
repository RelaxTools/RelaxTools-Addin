VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CaptureScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------------------
'
' [RelaxTools-Addin] v4
'
' Copyright (c) 2009 Yasuhiro Watanabe
' https://github.com/RelaxTools/RelaxTools-Addin
' author:relaxtools@opensquare.net
'
' The MIT License (MIT)
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'
'-----------------------------------------------------------------------------------------------------
Option Explicit

#If VBA7 And Win64 Then

    Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
    
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function BitBlt Lib "gdi32" (ByVal hDestDC As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
    
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr

    Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As rect) As Long
    Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "User32.dll" (ByVal wFormat As Long) As Long
    Private Declare PtrSafe Function GetWindowDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As LongPtr) As LongPtr
    Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
    Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
    Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
    
    Private Declare PtrSafe Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As LongPtr, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As LongPtr
    Private Declare PtrSafe Function PrintWindow Lib "User32.dll" (ByVal hwnd As LongPtr, ByVal hdcBlt As LongPtr, ByVal nFlags As Long) As Long
    Private Declare PtrSafe Function UpdateWindow Lib "user32" (ByVal hwnd As LongPtr) As Long

    Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As LongPtr, IPic As IPicture) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As LongPtr, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
    Private Declare PtrSafe Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As LongPtr
    Private Declare PtrSafe Function SelectPalette Lib "gdi32" (ByVal hdc As LongPtr, ByVal hPalette As LongPtr, ByVal bForceBackground As Long) As LongPtr
    Private Declare PtrSafe Function RealizePalette Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

    Private Declare PtrSafe Function DwmGetWindowAttribute Lib "Dwmapi" (ByVal hwnd As LongPtr, ByVal t As Long, lpRect As rect, ByVal Size As Long) As Long
    
    Private Declare PtrSafe Function PatBlt Lib "gdi32" (ByVal hdc As LongPtr, ByVal nXLeft As Long, ByVal nYLeft As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
    
#Else

    Private Declare Function GetDesktopWindow Lib "user32" () As Long
    Private Declare Function GetForegroundWindow Lib "user32" () As Long
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Private Declare Function BitBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, ByVal hdcS As Long, ByVal xS As Long, ByVal yS As Long, ByVal dwRop As Long) As Long
    
    Private Declare Function OpenClipboard Lib "user32" (ByVal hWndNewOwner As Long) As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function EmptyClipboard Lib "user32" () As Long
    Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As rect) As Long
    Private Declare Function SetClipboardData Lib "user32" (ByVal uFormat As Long, ByVal hMem As Long) As Long
    
    Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As rect) As Long
    Private Declare Function IsClipboardFormatAvailable Lib "User32.dll" (ByVal wFormat As Long) As Long
    Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
    Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
    Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hgdiobj As Long) As Long
    Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    
    Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpszName As String, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
    Private Declare Function PrintWindow Lib "User32.dll" (ByVal hwnd As Long, ByVal hdcBlt As Long, ByVal nFlags As Long) As Long
    Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As LongPtr) As Long

    Private Declare Function OleCreatePictureIndirect Lib "olepro32" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal iCapabilitiy As Long) As Long
    Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
    Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
    Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
    Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

    Private Declare Function DwmGetWindowAttribute Lib "Dwmapi" (ByVal hwnd As LongPtr, ByVal t As Long, lpRect As rect, ByVal Size As Long) As Long
    Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal nXLeft As Long, ByVal nYLeft As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
    Private Declare Function DwmIsCompositionEnabled Lib "Dwmapi" (b As Boolean) As Long
    

#End If


#If VBA7 And Win64 Then
    Private Type PicBmp
       Size As Long
       Type As Long
       hBmp As LongPtr
       hPal As LongPtr
       Reserved As Long
    End Type
#Else
    Private Type PicBmp
       Size As Long
       Type As Long
       hBmp As Long
       hPal As Long
       Reserved As Long
    End Type
#End If

Private Type rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
End Type

Private Type LOGPALETTE
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(255) As PALETTEENTRY
End Type

Private Const SRCCOPY = &HCC0020
Private Const CAPTUREBLT = &H40000000
Private Const CF_BITMAP = 2
Private Const IMAGE_BITMAP = 0
Private Const LR_LOADFROMFILE = &H10
Private Const RASTERCAPS As Long = 38
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104

Private Const SM_CXSCREEN As Long = 0
Private Const SM_CYSCREEN As Long = 1

Private Const WHITENESS As Long = &HFF0062
#If VBA7 And Win64 Then
    Private Function CaptureWindow(ByVal hWndSrc As LongPtr, ByVal Client As Boolean, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As IPictureDisp
        
        Dim r As Long
        
        Dim hDCSrc As LongPtr
        Dim hDCMemory As LongPtr
        Dim hBmp As LongPtr
        Dim hBmpPrev As LongPtr
        Dim RasterCapsScrn As Long
        Dim HasPaletteScrn As Long
        Dim PaletteSizeScrn As Long
        Dim hPal As LongPtr
        Dim hPalPrev As LongPtr
        Dim LogPal As LOGPALETTE
        
        hDCSrc = IIf(Client, GetDC(hWndSrc), GetWindowDC(hWndSrc))
        
        hDCMemory = CreateCompatibleDC(hDCSrc)
        
        hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
        hBmpPrev = SelectObject(hDCMemory, hBmp)
        
        RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS)
        HasPaletteScrn = RasterCapsScrn And RC_PALETTE
        PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE)
        
        If HasPaletteScrn And (PaletteSizeScrn = 256) Then
            LogPal.palVersion = &H300
            LogPal.palNumEntries = 256
            r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
            hPal = CreatePalette(LogPal)
            hPalPrev = SelectPalette(hDCMemory, hPal, 0)
            r = RealizePalette(hDCMemory)
        End If
        
        r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, SRCCOPY)
        
        hBmp = SelectObject(hDCMemory, hBmpPrev)
        
        If HasPaletteScrn And (PaletteSizeScrn = 256) Then
            hPal = SelectPalette(hDCMemory, hPalPrev, 0)
        End If
        
        r = DeleteDC(hDCMemory)
        r = ReleaseDC(hWndSrc, hDCSrc)
        
        Dim IID_IDispatch As GUID
        With IID_IDispatch
            .Data1 = &H20400
            .Data4(0) = &HC0
            .Data4(7) = &H46
        End With
        
        Dim pic As PicBmp
        With pic
            .Size = Len(pic)
            .Type = 1
            .hBmp = hBmp
            .hPal = hPal
        End With
        
        Dim IPic As IPicture
        
        r = OleCreatePictureIndirect(pic, IID_IDispatch, 1, IPic)
        
        Set CaptureWindow = IPic
    
    End Function
#Else
    Private Function CaptureWindow(ByVal hWndSrc As Long, ByVal Client As Boolean, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As IPictureDisp
        
        Dim r As Long
        
        Dim hDCSrc As Long
        Dim hDCMemory As Long
        Dim hBmp As Long
        Dim hBmpPrev As Long
        Dim RasterCapsScrn As Long
        Dim HasPaletteScrn As Long
        Dim PaletteSizeScrn As Long
        Dim hPal As Long
        Dim hPalPrev As Long
        Dim LogPal As LOGPALETTE
        
        hDCSrc = IIf(Client, GetDC(hWndSrc), GetWindowDC(hWndSrc))
        
        hDCMemory = CreateCompatibleDC(hDCSrc)
        
        hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
        hBmpPrev = SelectObject(hDCMemory, hBmp)
        
'        r = PatBlt(hBmp, 0, 0, WidthSrc, HeightSrc, WHITENESS)
        
        RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS)
        HasPaletteScrn = RasterCapsScrn And RC_PALETTE
        PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE)
        
        If HasPaletteScrn And (PaletteSizeScrn = 256) Then
            LogPal.palVersion = &H300
            LogPal.palNumEntries = 256
            r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
            hPal = CreatePalette(LogPal)
            hPalPrev = SelectPalette(hDCMemory, hPal, 0)
            r = RealizePalette(hDCMemory)
        End If
        
        r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, SRCCOPY)
        
        hBmp = SelectObject(hDCMemory, hBmpPrev)
        
        If HasPaletteScrn And (PaletteSizeScrn = 256) Then
            hPal = SelectPalette(hDCMemory, hPalPrev, 0)
        End If
        
        r = DeleteDC(hDCMemory)
        r = ReleaseDC(hWndSrc, hDCSrc)
        
        Dim IID_IDispatch As GUID
        With IID_IDispatch
            .Data1 = &H20400
            .Data4(0) = &HC0
            .Data4(7) = &H46
        End With
        
        Dim pic As PicBmp
        With pic
            .Size = Len(pic)
            .Type = 1
            .hBmp = hBmp
            .hPal = hPal
        End With
        
        Dim IPic As IPicture
        
        r = OleCreatePictureIndirect(pic, IID_IDispatch, 1, IPic)
        
        Set CaptureWindow = IPic
    
    End Function
#End If
#If VBA7 And Win64 Then
    Private Sub ClipCaptureWindow(ByVal hWndSrc As LongPtr, ByVal Client As Boolean, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long)
        
        Dim r As Long
        
        Dim hDCSrc As LongPtr
        Dim hDCMemory As LongPtr
        Dim hBmp As LongPtr
        Dim hBmpPrev As LongPtr
        Dim RasterCapsScrn As Long
        Dim HasPaletteScrn As Long
        Dim PaletteSizeScrn As Long
        Dim hPal As LongPtr
        Dim hPalPrev As LongPtr
        Dim LogPal As LOGPALETTE
        
        hDCSrc = IIf(Client, GetDC(hWndSrc), GetWindowDC(hWndSrc))
        
        hDCMemory = CreateCompatibleDC(hDCSrc)
        
        hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
        hBmpPrev = SelectObject(hDCMemory, hBmp)
        
        RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS)
        HasPaletteScrn = RasterCapsScrn And RC_PALETTE
        PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE)
        
        If HasPaletteScrn And (PaletteSizeScrn = 256) Then
            LogPal.palVersion = &H300
            LogPal.palNumEntries = 256
            r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
            hPal = CreatePalette(LogPal)
            hPalPrev = SelectPalette(hDCMemory, hPal, 0)
            r = RealizePalette(hDCMemory)
        End If
        
        r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, SRCCOPY)
        
        hBmp = SelectObject(hDCMemory, hBmpPrev)
        
        If HasPaletteScrn And (PaletteSizeScrn = 256) Then
            hPal = SelectPalette(hDCMemory, hPalPrev, 0)
        End If
        
        r = DeleteDC(hDCMemory)
        r = ReleaseDC(hWndSrc, hDCSrc)
        
        If OpenClipboard(0) Then
            Call EmptyClipboard
            Call SetClipboardData(CF_BITMAP, hBmp)
            Call CloseClipboard
         End If
    
    End Sub
#Else
    Private Sub ClipCaptureWindow(ByVal hWndSrc As Long, ByVal Client As Boolean, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long)
        
        Dim r As Long
        
        Dim hDCSrc As Long
        Dim hDCMemory As Long
        Dim hBmp As Long
        Dim hBmpPrev As Long
        Dim RasterCapsScrn As Long
        Dim HasPaletteScrn As Long
        Dim PaletteSizeScrn As Long
        Dim hPal As Long
        Dim hPalPrev As Long
        Dim LogPal As LOGPALETTE
        
        hDCSrc = IIf(Client, GetDC(hWndSrc), GetWindowDC(hWndSrc))
        
        hDCMemory = CreateCompatibleDC(hDCSrc)
        
        hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
        hBmpPrev = SelectObject(hDCMemory, hBmp)
        
        RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS)
        HasPaletteScrn = RasterCapsScrn And RC_PALETTE
        PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE)
        
        If HasPaletteScrn And (PaletteSizeScrn = 256) Then
            LogPal.palVersion = &H300
            LogPal.palNumEntries = 256
            r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
            hPal = CreatePalette(LogPal)
            hPalPrev = SelectPalette(hDCMemory, hPal, 0)
            r = RealizePalette(hDCMemory)
        End If
        
        r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, SRCCOPY)
        
        hBmp = SelectObject(hDCMemory, hBmpPrev)
        
        If HasPaletteScrn And (PaletteSizeScrn = 256) Then
            hPal = SelectPalette(hDCMemory, hPalPrev, 0)
        End If
        
        r = DeleteDC(hDCMemory)
        r = ReleaseDC(hWndSrc, hDCSrc)
        
        If OpenClipboard(0) Then
            Call EmptyClipboard
            Call SetClipboardData(CF_BITMAP, hBmp)
            Call CloseClipboard
         End If
    
    End Sub
#End If
Public Function GetPictureFromCaptureScreen() As IPictureDisp

    #If VBA7 And Win64 Then
        Dim hWndScreen As LongPtr
    #Else
        Dim hWndScreen As Long
    #End If
    
    hWndScreen = GetDesktopWindow()
    
    Dim cxScreen As Long, cyScreen As Long
    cxScreen = GetSystemMetrics(SM_CXSCREEN)
    cyScreen = GetSystemMetrics(SM_CYSCREEN)
    
    Set GetPictureFromCaptureScreen = CaptureWindow(hWndScreen, False, 0, 0, cxScreen, cyScreen)
    
End Function
Public Function GetPictureFromCaptureActiveWindow() As IPictureDisp

    Dim r As Long

    #If VBA7 And Win64 Then
        Dim hWndScreen As LongPtr
    #Else
        Dim hWndScreen As Long
    #End If
    
    hWndScreen = GetForegroundWindow()

    Dim RectActive As rect
    r = GetWindowRect(hWndScreen, RectActive)

    Set GetPictureFromCaptureActiveWindow = CaptureWindow(hWndScreen, False, 0, 0, RectActive.Right - RectActive.Left, RectActive.Bottom - RectActive.Top)

End Function
Public Sub SaveFileFromCaptureScreen(ByVal strFile As String)

    #If VBA7 And Win64 Then
        Dim hWndScreen As LongPtr
    #Else
        Dim hWndScreen As Long
    #End If
    
    hWndScreen = GetDesktopWindow()
    
    Dim cxScreen As Long
    Dim cyScreen As Long
    
    cxScreen = GetSystemMetrics(SM_CXSCREEN)
    cyScreen = GetSystemMetrics(SM_CYSCREEN)
    
    SavePicture CaptureWindow(hWndScreen, False, 0, 0, cxScreen, cyScreen), strFile
    
End Sub
Public Sub SaveFileFromCaptureActiveWindow(ByVal strFile As String)

    Dim r As Long

    #If VBA7 And Win64 Then
        Dim hWndScreen As LongPtr
    #Else
        Dim hWndScreen As Long
    #End If
    
    hWndScreen = GetForegroundWindow()

    Dim RectActive As rect
'    Const DWMWA_EXTENDED_FRAME_BOUNDS = 9
'
'    Dim bEnable As Boolean
'    DwmIsCompositionEnabled bEnable
    
'    If bEnable Then
'        DwmGetWindowAttribute hWndScreen, DWMWA_EXTENDED_FRAME_BOUNDS, RectActive, Len(RectActive)
'    Else
        r = GetWindowRect(hWndScreen, RectActive)
'    End If
    

    SavePicture CaptureWindow(hWndScreen, False, 0, 0, RectActive.Right - RectActive.Left, RectActive.Bottom - RectActive.Top), strFile

End Sub
Public Sub CopyClipboardFromCaptureScreen()

    #If VBA7 And Win64 Then
        Dim hWndScreen As LongPtr
    #Else
        Dim hWndScreen As Long
    #End If
    
    hWndScreen = GetDesktopWindow()
    
    Dim cxScreen As Long
    Dim cyScreen As Long
    
    cxScreen = GetSystemMetrics(SM_CXSCREEN)
    cyScreen = GetSystemMetrics(SM_CYSCREEN)
    
    Call ClipCaptureWindow(hWndScreen, False, 0, 0, cxScreen, cyScreen)
    
End Sub
Public Sub CopyClipboardFromCaptureActiveWindow()

    Dim r As Long

    #If VBA7 And Win64 Then
        Dim hWndScreen As LongPtr
    #Else
        Dim hWndScreen As Long
    #End If
    
    hWndScreen = GetForegroundWindow()

    Dim RectActive As rect
    r = GetWindowRect(hWndScreen, RectActive)
    
    Call ClipCaptureWindow(hWndScreen, False, 0, 0, RectActive.Right - RectActive.Left, RectActive.Bottom - RectActive.Top)

End Sub




