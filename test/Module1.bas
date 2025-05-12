Attribute VB_Name = "Module1"
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

'--- DIB Section constants
Private Const DIB_RGB_COLORS                As Long = 0
'--- for GetDeviceCaps
Private Const LOGPIXELSX                    As Long = 88
Private Const LOGPIXELSY                    As Long = 90
'--- for DrawIconEx
Private Const DI_NORMAL                     As Long = 3
'--- for CreateImagingFactory
Private Const WINCODEC_SDK_VERSION1         As Long = &H236
Private Const WINCODEC_SDK_VERSION2         As Long = &H237
'--- for IWICBitmapScaler
Private Const WICBitmapInterpolationModeFant As Long = 3
Private Const WICBitmapInterpolationModeHighQualityCubic As Long = 4

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, lpBitsInfo As BITMAPINFOHEADER, ByVal wUsage As Long, lpBits As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, pIconInfo As ICONINFO) As Long
Private Declare Function APIXbrzScale Lib "xbrz" Alias "XbrzScale" (ByVal lFactor As Long, ByVal lpSrc As Long, ByVal lpDst As Long, ByVal lWidth As Long, ByVal lHeight As Long, ByVal eFormat As XbrzColorFormat) As Long
Private Declare Function APIXbrzBilinearScale Lib "xbrz" Alias "XbrzBilinearScale" (ByVal lpSrc As Long, ByVal lSrcWidth As Long, ByVal lSrcHeight As Long, ByVal lpDst As Long, ByVal lDstWidth As Long, ByVal lDstHeight As Long) As Long
Private Declare Function CreateIconIndirect Lib "user32" (pIconInfo As ICONINFO) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "oleaut32" (lpPictDesc As PICTDESC, riid As Any, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
'--- WIC
Private Declare Function WICCreateImagingFactory_Proxy Lib "windowscodecs" (ByVal SDKVersion As Long, ppIImagingFactory As stdole.IUnknown) As Long
Private Declare Function IWICImagingFactory_CreateBitmapFromMemory_Proxy Lib "windowscodecs" (ByVal this As stdole.IUnknown, ByVal uiWidth As Long, ByVal uiHeight As Long, PixelFormat As Any, ByVal cbStride As Long, ByVal cbBufferSize As Long, pbBuffer As Any, ppIBitmap As stdole.IUnknown) As Long
Private Declare Function IWICImagingFactory_CreateBitmapScaler_Proxy Lib "windowscodecs" (ByVal pFactory As stdole.IUnknown, ppIBitmapScaler As stdole.IUnknown) As Long
Private Declare Function IWICBitmapScaler_Initialize_Proxy Lib "windowscodecs" (ByVal pThis As stdole.IUnknown, ByVal pISource As stdole.IUnknown, ByVal uiWidth As Long, ByVal uiHeight As Long, ByVal lMode As Long) As Long
Private Declare Function IWICBitmapSource_CopyPixels_Proxy Lib "windowscodecs" (ByVal pThis As stdole.IUnknown, prc As Any, ByVal cbStride As Long, ByVal cbBufferSize As Long, pbBuffer As Any) As Long

Private Type BITMAPINFOHEADER
    biSize              As Long
    biWidth             As Long
    biHeight            As Long
    biPlanes            As Integer
    biBitCount          As Integer
    biCompression       As Long
    biSizeImage         As Long
    biXPelsPerMeter     As Long
    biYPelsPerMeter     As Long
    biClrUsed           As Long
    biClrImportant      As Long
End Type

Private Type ICONINFO
    fIcon               As Long
    xHotspot            As Long
    yHotspot            As Long
    hbmMask             As Long
    hbmColor            As Long
End Type

Private Type PICTDESC
    lSize               As Long
    lType               As Long
    hBmp                As Long
    hPal                As Long
End Type

Public Enum XbrzColorFormat
    XbrzColorFormat_RGB                 ' 8 bit for each red, green, blue, upper 8 bits unused
    XbrzColorFormat_ARGB                ' including alpha channel, BGRA byte order on little-endian machines
    XbrzColorFormat_ARGB_UNBUFFERED     ' like ARGB, but without the one-time buffer creation overhead (ca. 100 - 300 ms) at the expense of a slightly slower scaling time
End Enum

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_pWicFactory           As stdole.IUnknown
Private g_sngScreenTwipsPerPixelX As Single
Private g_sngScreenTwipsPerPixelY As Single

Public Property Get ScreenTwipsPerPixelX() As Single
    If g_sngScreenTwipsPerPixelX = 0 Then
        pvSetupScreenTwipsPerPixel
    End If
    ScreenTwipsPerPixelX = g_sngScreenTwipsPerPixelX
End Property

Public Property Get ScreenTwipsPerPixelY() As Single
    If g_sngScreenTwipsPerPixelY = 0 Then
        pvSetupScreenTwipsPerPixel
    End If
    ScreenTwipsPerPixelY = g_sngScreenTwipsPerPixelY
End Property

Private Sub pvSetupScreenTwipsPerPixel()
    Dim hScreenDC       As Long
    
    hScreenDC = CreateCompatibleDC(0)
    g_sngScreenTwipsPerPixelX = GetDeviceCaps(hScreenDC, LOGPIXELSX)
    If g_sngScreenTwipsPerPixelX = 0 Then
        g_sngScreenTwipsPerPixelX = Screen.TwipsPerPixelX
    Else
        g_sngScreenTwipsPerPixelX = 1440# / g_sngScreenTwipsPerPixelX
    End If
    g_sngScreenTwipsPerPixelY = GetDeviceCaps(hScreenDC, LOGPIXELSY)
    If g_sngScreenTwipsPerPixelY = 0 Then
        g_sngScreenTwipsPerPixelY = Screen.TwipsPerPixelY
    Else
        g_sngScreenTwipsPerPixelY = 1440# / g_sngScreenTwipsPerPixelY
    End If
    Call DeleteDC(hScreenDC)
End Sub

Public Function IconScale(ByVal sngSize As Single) As Long
    If ScreenTwipsPerPixelX < 5.5 Then
        IconScale = Int(sngSize * 3)
    ElseIf ScreenTwipsPerPixelX < 6.7 Then
        IconScale = Int(sngSize * 2.5)
    ElseIf ScreenTwipsPerPixelX < 8.6 Then
        IconScale = Int(sngSize * 2)
    ElseIf ScreenTwipsPerPixelX < 12.1 Then
        IconScale = Int(sngSize * 1.5)
    Else
        IconScale = Int(sngSize * 1)
    End If
End Function

Public Function ScalePicture( _
            oPic As StdPicture, _
            Optional ByVal MaskColor As OLE_COLOR = -1, _
            Optional ByVal lTargetWidth As Long, _
            Optional ByVal lTargetHeight As Long, _
            Optional ByVal SkipXbrz As Boolean) As StdPicture
    Dim lWidth          As Long
    Dim lHeight         As Long
    Dim aSrcBits()      As Long
    Dim aDstBits()      As Long
    Dim bHasAlpha       As Boolean
    Dim lFactor         As Long
    Dim hMemDC          As Long
    Dim uHdr            As BITMAPINFOHEADER
    Dim hDib            As Long
    Dim lpBits          As Long
    Dim uInfo           As ICONINFO
    Dim hIcon           As Long
    Dim uDesc           As PICTDESC
    Dim aGUID(0 To 3)   As Long
    
    Set ScalePicture = oPic
    If oPic Is Nothing Then
        GoTo QH
    End If
    If oPic.Handle = 0 Then
        GoTo QH
    End If
    If MaskColor <> -1 Then
        MaskColor = TranslateColor(MaskColor)
    End If
    If Not pvGetDIBits(oPic, MaskColor, aSrcBits, lWidth, lHeight, bHasAlpha) Then
        GoTo QH
    End If
    If lTargetWidth = 0 Or lTargetHeight = 0 Then
        lTargetWidth = IconScale(lWidth)
        lTargetHeight = IconScale(lHeight)
    End If
    lFactor = Clamp(Ceil(lTargetWidth / lWidth), 2, 6)
    If SkipXbrz Then
        GoTo DoResize
    End If
    ReDim aDstBits(0 To lWidth * lFactor * lHeight * lFactor - 1) As Long
    If XbrzScale(lFactor, VarPtr(aSrcBits(0)), VarPtr(aDstBits(0)), lWidth, lHeight, IIf(bHasAlpha, XbrzColorFormat_ARGB, XbrzColorFormat_RGB)) = 0 Then
        GoTo DoResize
    End If
    lWidth = lWidth * lFactor
    lHeight = lHeight * lFactor
    If lTargetWidth <> lWidth Or lTargetHeight <> lHeight Then
        aSrcBits = aDstBits
DoResize:
        ReDim aDstBits(0 To lTargetWidth * lTargetHeight - 1) As Long
        If Not WicBicubicScale(VarPtr(aSrcBits(0)), lWidth, lHeight, VarPtr(aDstBits(0)), lTargetWidth, lTargetHeight, bHasAlpha) Then
            If XbrzBilinearScale(VarPtr(aSrcBits(0)), lWidth, lHeight, VarPtr(aDstBits(0)), lTargetWidth, lTargetHeight) = 0 Then
                GoTo QH
            End If
        End If
    End If
    hMemDC = CreateCompatibleDC(0)
    If hMemDC = 0 Then
        GoTo QH
    End If
    With uHdr
        .biSize = LenB(uHdr)
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = lTargetWidth
        .biHeight = -lTargetHeight
        .biSizeImage = 4 * lTargetWidth * lTargetHeight
    End With
    hDib = CreateDIBSection(hMemDC, uHdr, DIB_RGB_COLORS, lpBits, 0, 0)
    If hDib = 0 Then
        GoTo QH
    End If
    Debug.Assert uHdr.biSizeImage = 4 * (UBound(aDstBits) + 1)
    Call CopyMemory(ByVal lpBits, aDstBits(0), uHdr.biSizeImage)
    With uInfo
        .fIcon = 1
        .hbmColor = hDib
        .hbmMask = hDib
    End With
    hIcon = CreateIconIndirect(uInfo)
    With uDesc
        .lSize = LenB(uDesc)
        .lType = vbPicTypeIcon
        .hBmp = hIcon
    End With
    aGUID(0) = &H7BF80980       '--- IID_IPicture = {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    If OleCreatePictureIndirect(uDesc, aGUID(0), 1, ScalePicture) < 0 Then
        GoTo QH
    End If
    hIcon = 0
QH:
    If hIcon <> 0 Then
        Call DestroyIcon(hIcon)
    End If
    If hDib <> 0 Then
        Call DeleteObject(hDib)
    End If
    If hMemDC <> 0 Then
        Call DeleteDC(hMemDC)
    End If
End Function

Private Function pvGetDIBits( _
            oPic As StdPicture, _
            clrMask As Long, _
            DIBits() As Long, _
            Optional lWidth As Long, _
            Optional lHeight As Long, _
            Optional bHasAlpha As Boolean) As Boolean
    Dim hMemDC          As Long
    Dim uInfo           As ICONINFO
    Dim uHdr            As BITMAPINFOHEADER
    Dim hDib            As Long
    Dim lpBits          As Long
    Dim hPrevDib        As Long
    Dim aMaskBits()     As Long
    Dim lIdx            As Long
    Dim pPic            As IPicture
    
    lWidth = HM2Pix(oPic.Width)
    lHeight = HM2Pix(oPic.Height)
    hMemDC = CreateCompatibleDC(0)
    If hMemDC = 0 Then
        GoTo QH
    End If
    With uHdr
        .biSize = LenB(uHdr)
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = lWidth
        .biHeight = -lHeight
        .biSizeImage = 4 * lWidth * lHeight
    End With
    If oPic.Type = vbPicTypeIcon Then
        If GetIconInfo(oPic.Handle, uInfo) = 0 Then
            GoTo QH
        End If
        ReDim DIBits(0 To lWidth * lHeight - 1) As Long
        If GetDIBits(hMemDC, uInfo.hbmColor, 0, lHeight, DIBits(0), uHdr, DIB_RGB_COLORS) = 0 Then
            GoTo QH
        End If
        ReDim aMaskBits(0 To lWidth * lHeight - 1) As Long
        If GetDIBits(hMemDC, uInfo.hbmMask, 0, lHeight, aMaskBits(0), uHdr, DIB_RGB_COLORS) = 0 Then
            GoTo QH
        End If
        For lIdx = 0 To UBound(aMaskBits)
            If aMaskBits(lIdx) = 0 Then
                DIBits(lIdx) = DIBits(lIdx) Or &HFF000000
            Else
                DIBits(lIdx) = 0
            End If
        Next
        bHasAlpha = True
    Else
        hDib = CreateDIBSection(hMemDC, uHdr, DIB_RGB_COLORS, lpBits, 0, 0)
        If hDib = 0 Then
            GoTo QH
        End If
        hPrevDib = SelectObject(hMemDC, hDib)
        Set pPic = oPic
        pPic.Render hMemDC, 0, 0, lWidth, lHeight, 0, pPic.Height, pPic.Width, -pPic.Height, ByVal 0
        ReDim DIBits(0 To lWidth * lHeight - 1) As Long
        Call CopyMemory(DIBits(0), ByVal lpBits, uHdr.biSizeImage)
        bHasAlpha = False
        For lIdx = 0 To UBound(DIBits)
            If (DIBits(lIdx) And &HFF000000) <> 0 Then
                bHasAlpha = True
                Exit For
            End If
        Next
        If Not bHasAlpha And clrMask <> -1 Then
            bHasAlpha = True
            For lIdx = 0 To UBound(DIBits)
                If DIBits(lIdx) = clrMask Then
                    DIBits(lIdx) = 0
                Else
                    DIBits(lIdx) = DIBits(lIdx) Or &HFF000000
                End If
            Next
        End If
    End If
    '--- success
    pvGetDIBits = True
QH:
    If hPrevDib <> 0 Then
        Call SelectObject(hMemDC, hPrevDib)
    End If
    If hDib <> 0 Then
        Call DeleteObject(hDib)
    End If
    If uInfo.hbmColor <> 0 Then
        Call DeleteObject(uInfo.hbmColor)
    End If
    If uInfo.hbmMask <> 0 Then
        Call DeleteObject(uInfo.hbmMask)
    End If
    If hMemDC <> 0 Then
        Call DeleteDC(hMemDC)
    End If
End Function

Private Function XbrzScale(ByVal lFactor As Long, ByVal lpSrc As Long, ByVal lpDst As Long, ByVal lWidth As Long, ByVal lHeight As Long, ByVal eFormat As XbrzColorFormat) As Long
    Debug.Assert lpSrc <> 0 And lpDst <> 0
    On Error GoTo QH
    XbrzScale = APIXbrzScale(lFactor, lpSrc, lpDst, lWidth, lHeight, eFormat)
QH:
End Function

Private Function XbrzBilinearScale(ByVal lpSrc As Long, ByVal lSrcWidth As Long, ByVal lSrcHeight As Long, ByVal lpDst As Long, ByVal lDstWidth As Long, ByVal lDstHeight As Long) As Long
    Debug.Assert lpSrc <> 0 And lpDst <> 0
    On Error GoTo QH
    XbrzBilinearScale = APIXbrzBilinearScale(lpSrc, lSrcWidth, lSrcHeight, lpDst, lDstWidth, lDstHeight)
QH:
End Function

Private Function WicBicubicScale(ByVal lpSrc As Long, ByVal lSrcWidth As Long, ByVal lSrcHeight As Long, ByVal lpDst As Long, ByVal lDstWidth As Long, ByVal lDstHeight As Long, Optional ByVal HasAlpha As Long) As Boolean
    Dim aGUID(0 To 3)   As Long
    Dim pBitmap         As stdole.IUnknown
    Dim pScaler         As stdole.IUnknown
    
    Debug.Assert lpSrc <> 0 And lpDst <> 0
    If m_pWicFactory Is Nothing Then
        If WICCreateImagingFactory_Proxy(WINCODEC_SDK_VERSION2, m_pWicFactory) < 0 Then
            If WICCreateImagingFactory_Proxy(WINCODEC_SDK_VERSION1, m_pWicFactory) < 0 Then
                GoTo QH
            End If
        End If
    End If
    aGUID(0) = &H6FDDC324       ' GUID_WICPixelFormat32bppBGR = {6FDDC324-4E03-4BFE-B185-3D77768DC90E}
    aGUID(1) = &H4BFE4E03
    aGUID(2) = &H773D85B1
    aGUID(3) = &HEC98D76
    If HasAlpha = 2 Then
        aGUID(3) = &H10C98D76   ' GUID_WICPixelFormat32bppPBGRA = {6FDDC324-4E03-4BFE-B185-3D77768DC910}
    ElseIf HasAlpha <> 0 Then
        aGUID(3) = &HFC98D76    ' GUID_WICPixelFormat32bppBGRA = {6FDDC324-4E03-4BFE-B185-3D77768DC90F}
    End If
    If IWICImagingFactory_CreateBitmapFromMemory_Proxy(m_pWicFactory, lSrcWidth, lSrcHeight, aGUID(0), 4 * lSrcWidth, 4 * lSrcWidth * lSrcHeight, ByVal lpSrc, pBitmap) < 0 Then
        GoTo QH
    End If
    If IWICImagingFactory_CreateBitmapScaler_Proxy(m_pWicFactory, pScaler) < 0 Then
        GoTo QH
    End If
    If IWICBitmapScaler_Initialize_Proxy(pScaler, pBitmap, lDstWidth, lDstHeight, WICBitmapInterpolationModeHighQualityCubic) < 0 Then
        If IWICBitmapScaler_Initialize_Proxy(pScaler, pBitmap, lDstWidth, lDstHeight, WICBitmapInterpolationModeFant) < 0 Then
            GoTo QH
        End If
    End If
    If IWICBitmapSource_CopyPixels_Proxy(pScaler, ByVal 0, 4 * lDstWidth, 4 * lDstWidth * lDstHeight, ByVal lpDst) < 0 Then
        GoTo QH
    End If
    WicBicubicScale = True
QH:
End Function

'= shared ================================================================

Public Function TranslateColor(ByVal clrValue As OLE_COLOR) As OLE_COLOR
    Call OleTranslateColor(clrValue, 0, VarPtr(TranslateColor))
End Function

Public Function HM2Pix(ByVal Value As Double) As Long
    HM2Pix = Int(Value * 1440 / 2540 / ScreenTwipsPerPixelX + 0.5)
End Function

Public Function Ceil(ByVal Value As Double) As Double
    Ceil = -Int(-Value)
End Function

Public Function Clamp( _
            ByVal lValue As Long, _
            Optional ByVal Min As Long = -2147483647, _
            Optional ByVal Max As Long = 2147483647) As Long
    If lValue < Min Then
        Clamp = Min
    ElseIf lValue > Max Then
        Clamp = Max
    Else
        Clamp = lValue
    End If
End Function

Public Sub RenderPicture(pPic As IPicture, hDC As Long, X As Long, Y As Long, cx As Long, cy As Long, xSrc As OLE_XPOS_HIMETRIC, ySrc As OLE_YPOS_HIMETRIC, cxSrc As OLE_XSIZE_HIMETRIC, cySrc As OLE_YSIZE_HIMETRIC)
    If pPic Is Nothing Then
        Exit Sub
    End If
    If pPic.Handle = 0 Then
        Exit Sub
    End If
    If pPic.Type = vbPicTypeIcon Then
        Call DrawIconEx(hDC, X, Y, pPic.Handle, cx, cy, 0, 0, DI_NORMAL)
    Else
        pPic.Render hDC, X, Y, cx, cy, xSrc, ySrc, cxSrc, cySrc, ByVal 0
    End If
End Sub
