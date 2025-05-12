VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "xBRZ upscale"
   ClientHeight    =   11220
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11976
   LinkTopic       =   "Form1"
   ScaleHeight     =   935
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   998
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPreview 
      Align           =   4  'Align Right
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   11220
      Left            =   5916
      ScaleHeight     =   935
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   505
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6060
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'=========================================================================
' API
'=========================================================================

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExW" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As Long, ByVal lpszWindow As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropW" (ByVal hWnd As Long, ByVal lpString As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private m_bActive As Boolean

Private Type UcsHitInfo
    Left            As Long
    Top             As Long
    Right           As Long
    Bottom          As Long
    Picture         As StdPicture
End Type
    
Private m_uHitInfo()    As UcsHitInfo

Private Property Get VbeRef() As VBIDE.VBE
    Static oRetVal  As VBIDE.VBE
    Dim hWnd        As Long
    Dim lProcessId  As Long
    Dim hProp       As Long
    Dim oWindow     As VBIDE.Window
    
    If oRetVal Is Nothing Then
        Do
            hWnd = FindWindowEx(0, hWnd, StrPtr("wndclass_desked_gsk"), 0)
            Call GetWindowThreadProcessId(hWnd, lProcessId)
        Loop While hWnd <> 0 And lProcessId <> GetCurrentProcessId()
        hProp = GetProp(hWnd, StrPtr("VBAutomation"))
        If hProp <> 0 Then
            Call vbaObjSetAddref(oWindow, hProp)
            Set oRetVal = oWindow.VBE
        End If
    End If
    Set VbeRef = oRetVal
End Property

Private Sub pvClipClear()
    Dim lRetry          As Long
    
    On Error GoTo EH
    Clipboard.Clear
    Exit Sub
EH:
    If lRetry < 100 Then
        lRetry = lRetry + 1
        Call Sleep(1)
        Resume
    End If
End Sub

Private Function pvClipGetData() As IPictureDisp
    Dim lRetry          As Long
    
    On Error GoTo EH
    Set pvClipGetData = Clipboard.GetData()
    Exit Function
EH:
    If lRetry < 100 Then
        lRetry = lRetry + 1
        Call Sleep(1)
        Resume
    End If
End Function

Private Sub pvBtnCopyFace(oBtn As Object)
    Dim lRetry          As Long
    
    On Error GoTo EH
    oBtn.CopyFace
    Exit Sub
EH:
    If lRetry < 100 Then
        lRetry = lRetry + 1
        Call Sleep(1)
        Resume
    End If
End Sub

Private Sub pvRenderUpscaled(oBox As PictureBox, oSrc As StdPicture)
    Dim vElem           As Variant
    Dim oPic            As StdPicture
    
    oBox.Cls
    oBox.CurrentX = 50
    oBox.CurrentY = 0
    oBox.Print "xBRZ"
    oBox.CurrentX = 50 + 150
    oBox.CurrentY = 0
    oBox.Print "Bicubic"
    oBox.CurrentX = 50 + 150 + 150
    oBox.CurrentY = 0
    oBox.Print "Stretched"
    oBox.CurrentY = 32
    For Each vElem In Array(Empty, 24, 32, 40, 48, 64, 80, 96)
        oBox.CurrentX = 6
        If Not IsEmpty(vElem) Then
            oBox.Print vElem & " px"
        Else
            oBox.Print "Orig."
            Set oPic = oSrc
        End If
        oBox.CurrentX = 50
        '--- xBRZ upscale (w/ bicubic downsample)
        If Not IsEmpty(vElem) Then
            Set oPic = ScalePicture(oSrc, vbButtonFace, vElem, vElem)
        End If
        RenderPicture oPic, oBox.hDC, oBox.CurrentX, oBox.CurrentY - 16, HM2Pix(oPic.Width), HM2Pix(oPic.Height), 0, oPic.Height, oPic.Width, -oPic.Height
        oBox.CurrentX = oBox.CurrentX + 150
        '--- bicubic resize
        If Not IsEmpty(vElem) Then
            Set oPic = ScalePicture(oSrc, vbButtonFace, vElem, vElem, SkipXbrz:=True)
        End If
        RenderPicture oPic, oBox.hDC, oBox.CurrentX, oBox.CurrentY - 16, HM2Pix(oPic.Width), HM2Pix(oPic.Height), 0, oPic.Height, oPic.Width, -oPic.Height
        oBox.CurrentX = oBox.CurrentX + 150
        '--- nearest neighbor
        If Not IsEmpty(vElem) Then
            RenderPicture oSrc, oBox.hDC, oBox.CurrentX, oBox.CurrentY - 16, HM2Pix(oPic.Width), HM2Pix(oPic.Height), 0, oPic.Height * HM2Pix(oSrc.Height) / vElem, oPic.Width * HM2Pix(oSrc.Width) / vElem, -oPic.Height * HM2Pix(oSrc.Height) / vElem
        Else
            RenderPicture oPic, oBox.hDC, oBox.CurrentX, oBox.CurrentY - 16, HM2Pix(oPic.Width), HM2Pix(oPic.Height), 0, oPic.Height, oPic.Width, -oPic.Height
        End If
        oBox.CurrentY = oBox.CurrentY + HM2Pix(oPic.Height) + 16
        oBox.Refresh
    Next
End Sub

'=========================================================================
' Event handlers
'=========================================================================

Private Sub Form_Activate()
    Dim oBar        As Object
    Dim oBtn        As Object
    Dim pPic        As IPicture
    Dim lIdx        As Long
    
    If m_bActive Then
        Exit Sub
    End If
    m_bActive = True
    ReDim m_uHitInfo(0 To 1000) As UcsHitInfo
    For Each oBar In VbeRef.CommandBars
        If Left$(oBar.Name, 8) <> "DataView" And Left$(oBar.Name, 7) <> "Toolbox" And oBar.Name <> "Color Palette" And oBar.Name <> "Property Browser" Then
            If CurrentX <> 260 Then
                CurrentY = CurrentY + 16
                CurrentX = 16
                Print oBar.Name
                CurrentX = 260
            End If
            For Each oBtn In oBar.Controls
                If oBtn.Type = 1 Then
                    pvClipClear
                    pvBtnCopyFace oBtn
                    If Clipboard.GetFormat(vbCFBitmap) Then
                        Set pPic = pvClipGetData()
                        With m_uHitInfo(lIdx)
                            .Left = CurrentX
                            .Top = CurrentY - 16
                            .Right = .Left + HM2Pix(pPic.Width)
                            .Bottom = .Top + HM2Pix(pPic.Height)
                            Set .Picture = pPic
                            RenderPicture pPic, hDC, .Left, .Top, .Right - .Left, .Bottom - .Top, 0, pPic.Height, pPic.Width, -pPic.Height
                            CurrentX = .Right + 16
                        End With
                        lIdx = lIdx + 1
                        Refresh
                    End If
                End If
            Next
        End If
        If CurrentY > 96 Then
            Exit For
        End If
    Next
    ReDim Preserve m_uHitInfo(0 To lIdx) As UcsHitInfo
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lIdx            As Long
    
    For lIdx = 0 To UBound(m_uHitInfo)
        With m_uHitInfo(lIdx)
            If .Left <= X And X < .Right And .Top <= Y And Y < .Bottom Then
                pvRenderUpscaled picPreview, .Picture
                Exit For
            End If
        End With
    Next
End Sub
