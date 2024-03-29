VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMemDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'////////////////////////////////////////////////////////
'///                  Memory DC Class
'///                   (clsMemDC.cls)
'///_____________________________________________________
'/// Memory DC for flicker free drawing.
'///_____________________________________________________
'/// Last modification  : Ago/07/2000
'/// Last modified by   : Leontti R.
'/// Modification reason: Created
'/// Project: RamoSoft Component Suite ' I borrowed this code from a another project from myself
'/// Author: Leontti A. Ramos M. (leontti@leontti.net)
'/// RamoSoft de Mexico S.A. de C.V.
'////////////////////////////////////////////////////////
Option Explicit

Public BackColor As OLE_COLOR
Private m_hWorkDC As Long
Private m_hBmpOld As Long
Private m_hBmp As Long
Private m_lWidth As Long
Private m_lHeight As Long
Private m_iScaleMode As ScaleModeConstants
Private m_iMFactor As Single
Private m_lOldFont As Long
Private Const PIXELS_PER_INCHE = 96
Private Const PIXELS_PER_CENTIMETER = 37.8
Private Const PIXELS_PER_POINT = 4.8
Private Const PIXELS_PER_PIXEL = 1
Private Const TWIPS_PER_POINT = 72
Private Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Const ETO_CLIPPED = 4
Private Const ETO_OPAQUE = 2
' Bitmap copying related
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function LPtoDP Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Const HALFTONE = 4
Private Const SRCCOPY = &HCC0020
' Drawig related functions
Private Declare Function DrawTextAPI Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, ByVal HPALETTE As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
Private Const PS_SOLID = 0
Private Const PS_DOT = 2
Private Const DT_WORDBREAK = &H10
Private Const DT_VCENTER = &H4
Private Const DT_CENTER = &H1
Private Const BKMODE_OPAQUE = 2
Private Const BKMODE_TRANSPARENT = 1

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type
' Font creation
Private Const CLIP_DEFAULT_PRECIS = 0 '*
Private Const PROOF_QUALITY = 2 '*
Private Const DEFAULT_PITCH = 0 '*
Private Const ANSI_CHARSET = 0 '*
Private Const DEFAULT_CHARSET = 1 '*
Private Const OEM_CHARSET = 255 '*
' Font Families
Private Const FF_DONTCARE = 0 '*    '  Don't care or don't know.
Private Const FF_ROMAN = 16      '  Variable stroke width, serifed.
' Font Weights
Private Const FW_NORMAL = 400 '*
Private Const FW_BOLD = 700 '*

Private Const OUT_DEFAULT_PRECIS = 0 '*
Private Const OUT_TT_ONLY_PRECIS = 7 '*

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hDC As Long, ByVal nMapMode As Long) As Long

Private Const LOGPIXELSY = 90 '*        '  Logical pixels/inch in Y
Private Const MM_HIMETRIC = 3
Private Const MM_LOMETRIC = 2
' Bitmap info
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Friend Sub Dither(Optional hDC As Long)
    Dim LnStyle As Long
    Dim LhDC As Long
    Dim LhBrush As Long
    Dim LhOldBrush As Long
    Dim LoPattern As StdPicture
    
    If (hDC = 0) Then
        hDC = m_hWorkDC
    End If
    Set LoPattern = LoadResPicture(101, vbResBitmap)
    LhBrush = CreatePatternBrush(LoPattern.Handle)
    LhOldBrush = SelectObject(hDC, LhBrush)
    Call PatBlt(hDC, 0, 0, m_lWidth, m_lHeight, &HA000C9) '&HA0329) ' &HA000C9)
    LhOldBrush = SelectObject(hDC, LhOldBrush)
    Call DeleteObject(LhBrush)
End Sub


Public Sub DrawBox(ByVal lLeft As Long, ByVal lTop As Long, _
    ByVal lWidth As Long, ByVal lHeight As Long, Optional lLineWidth As Integer = 1, _
    Optional lFColor As OLE_COLOR, Optional lBColor As OLE_COLOR = -1)
    On Error GoTo ERR_H
    Dim LnTop As Long
    Dim LnLeft As Long
    Dim LnWidth As Long
    Dim LnHeight As Long
    Dim LrBox As RECT
    Dim LrPos As POINTAPI
    Dim LhPen As Long
    Dim LhOldPen As Long
'////////////////////////////////////////
'///      Coordenates Calculation
'////////////////////////////////////////
    With LrBox
        .Left = lLeft
        .Top = lTop
        .Right = lLeft + lWidth
        .Bottom = lTop + lHeight
    End With
'////////////////////////////////////////
'///         Background Drawing
'///_____________________________________
'/// If background is not transparent
'/// (<> -1) then draws it...
'////////////////////////////////////////
    If (lBColor <> -1) Then
        Dim LhBrush As Long
        
        LhBrush = CreateSolidBrush(lBColor)
        Call FillRect(m_hWorkDC, LrBox, LhBrush)
        Call DeleteObject(LhBrush)
    End If
'////////////////////////////////////////
'///           Border Drawing
'///_____________________________________
'/// If border is not transparent (<> -1)
'/// then draws it...
'////////////////////////////////////////
    If (lFColor <> -1) Then
        ' Creates a pen to draw
        If (lLineWidth = 0) Then
            ' Dotted line
            LhPen = CreatePen(PS_DOT, 1, lFColor)
        Else
            ' Regular line
            LhPen = CreatePen(PS_SOLID, lLineWidth, lFColor)
        End If
        If (LhPen <> 0) Then
            LhOldPen = SelectObject(m_hWorkDC, LhPen)
        End If
        ' Draws box lines (regulad Rect API call draws a non
        ' transparent rectangle, is not useful here...
        With LrBox
            MoveToEx m_hWorkDC, .Left, .Top, LrPos
            LineTo m_hWorkDC, .Right, .Top
            LineTo m_hWorkDC, .Right, .Bottom
            LineTo m_hWorkDC, .Left, .Bottom
            LineTo m_hWorkDC, .Left, .Top
        End With
        If LhPen Then
            Call SelectObject(m_hWorkDC, LhOldPen)
            Call DeleteObject(LhPen)
        End If
    End If
    Exit Sub
ERR_H:
    Me.RaiseErr Err.Number, "DrawBox"
End Sub

'Public Sub DrawFillRect(ByVal lX1 As Long, ByVal lY1 As Long, _
'    ByVal lX2 As Long, ByVal lY2 As Long, ByVal lColor As Long, _
'    Optional ByStep As Boolean = False)
'    On Error GoTo ERR_H
'    Dim LrBox As RECT
'    Dim oBrush As Long
'
'    With LrBox
'        .Left = lX1
'        .Top = lY1
'        If ByStep Then
'            .Right = (.Left + lX2)
'            .Bottom = (.Top + lY2)
'        Else
'            .Right = lX2
'            .Bottom = lY2
'        End If
'    End With
'    oBrush = CreateSolidBrush(lColor)
'    Call FillRect(m_hWorkDC, LrBox, oBrush)
'    DeleteObject (oBrush)
'    Exit Sub
'ERR_H:
'    Me.RaiseErr Err.Number, "DrawFillRect"
'End Sub
    
'Friend Sub DrawRect(ByVal lX1 As Long, ByVal lY1 As Long, _
'    ByVal lX2 As Long, ByVal lY2 As Long, ByVal lColor As Variant, _
'    Optional ByStep As Boolean = False)
'    On Error GoTo ERR_H
'    Dim LrBox As RECT
'    Dim LoPen As Long
'    Dim LhOldPen As Long
'
'    With LrBox
'        .Left = lX1
'        .Top = lY1
'        If ByStep Then
'            .Right = (.Left + lX2)
'            .Bottom = (.Top + lY2)
'        Else
'            .Right = lX2
'            .Bottom = lY2
'        End If
'    End With
'    LoPen = CreatePen(PS_SOLID, 1, lColor)
'    LhOldPen = SelectObject(m_hWorkDC, LoPen)
'    With LrBox
'        Call Rectangle(m_hWorkDC, .Left, .Top, .Right, .Bottom)
'    End With
'    Call SelectObject(m_hWorkDC, LhOldPen)
'    Call DeleteObject(LoPen)
'    Exit Sub
'ERR_H:
'    Me.RaiseErr Err.Number, "DrawRect"
'End Sub

Public Sub DrawLine(lX1 As Long, lY1 As Long, _
    lX2 As Long, lY2 As Long, _
    lColor As OLE_COLOR, Optional lWidth As Long = 1)
    On Error GoTo ERR_H
    Dim LrPos As POINTAPI
    Dim LhPen As Long
    Dim LhOldPen As Long
    
    If (lWidth = 0) Then
        LhPen = CreatePen(PS_DOT, 1, lColor)
    Else
        LhPen = CreatePen(PS_SOLID, lWidth, lColor)
    End If
    If LhPen Then
        LhOldPen = SelectObject(m_hWorkDC, LhPen)
    End If
    ' Draws box lines
    MoveToEx m_hWorkDC, lX1, lY1, LrPos
    LineTo m_hWorkDC, lX2, lY2
    If LhPen Then
        Call SelectObject(m_hWorkDC, LhOldPen)
        Call DeleteObject(LhPen)
    End If
    Exit Sub
ERR_H:
    Me.RaiseErr Err.Number, "DrawLine"
End Sub


Public Sub DrawText(ByVal sText As String, ByVal lLeft As Long, ByVal lTop As Long, _
    ByVal lWidth As Long, ByVal lHeight As Long, ByVal lFColor As OLE_COLOR, _
    ByVal lBColor As OLE_COLOR, ByVal lFlags As Long)
    On Error GoTo ERR_H
    Dim lBrush As Long
    Dim LnOldBMode As Long
    Dim LrBox As RECT
'////////////////////////////////////////
'///      Coordenates Calculation
'////////////////////////////////////////
    With LrBox
        .Left = lLeft
        .Top = lTop
        .Right = (.Left + lWidth)
        .Bottom = (.Top + lHeight)
    End With
'////////////////////////////////////////
'///       Defaults Resolution
'////////////////////////////////////////
    If (lFlags = 0) Then
        lFlags = (DT_WORDBREAK Or DT_VCENTER Or DT_CENTER)
    End If
    If (lFColor = -1) Then
        lFColor = vbBlack
    End If
'////////////////////////////////////////
'///         Background Drawing
'///_____________________________________
'/// If background is not transparent
'/// (<> -1) then draws it...
'////////////////////////////////////////
    If (lBColor <> -1) Then
        lBrush = CreateSolidBrush(lBColor)
        Call FillRect(m_hWorkDC, LrBox, lBrush)
        DeleteObject (lBrush)
    End If
'////////////////////////////////////////
'///           Text Drawing
'////////////////////////////////////////
    LnOldBMode = SetBkMode(m_hWorkDC, BKMODE_TRANSPARENT)
    Call SetTextColor(m_hWorkDC, lFColor)
    DrawTextAPI m_hWorkDC, sText, Len(sText), LrBox, lFlags
    Call SetBkMode(m_hWorkDC, LnOldBMode)
    Exit Sub
ERR_H:
    Me.RaiseErr Err.Number, "DrawText"
End Sub

Friend Sub RaiseErr(ByVal lErrNum As RSErrorCode, Optional sRoutineName As String, _
    Optional sDescription As String)
    RaiseError lErrNum, TypeName(Me), sRoutineName, sDescription, Erl
End Sub


Public Sub BlitImage(lDestDC As Long, lLeft As Long, lTop As Long, _
    lWidth As Long, lHeight As Long, Optional bNormal As Boolean = True)
    If ((m_lWidth = lWidth) And (m_lHeight = lHeight)) Then
        BitBlt lDestDC, lLeft, lTop, lWidth, lHeight, m_hWorkDC, 0, 0, SRCCOPY
    Else
        Dim LrPA As POINTAPI
        Dim LnOldMode As Long
        Dim LnOldPalette As Long
        'Copy to destination DC
        Call SetBrushOrgEx(lDestDC, 0, 0, LrPA)
        Call StretchBlt(lDestDC, lLeft, lTop, lWidth, lHeight, m_hWorkDC, _
             0, 0, m_lWidth, m_lHeight, SRCCOPY)
        If Not bNormal Then
            Dither lDestDC
        End If
    End If
End Sub

Public Sub DrawPicture(oPicture As StdPicture, lLeft As Long, lTop As Long, _
    lWidth As Long, lHeight As Long, Optional bKeepRatio As Boolean = True, _
    Optional bVCenter As Boolean = True, Optional bHCenter As Boolean = True)
    On Error GoTo ERR_H
    Dim LhDCTemp As Long
    Dim LhOldPalette As Long
    Dim LhOldBitmap As Long
    Dim LrPoint As POINTAPI
    Dim LrReal As RECT
    
#Const USE_PALETTE = True
#Const USE_DISPLAY = False
#Const USE_BITBLT = True
#Const USE_BKMODE = False
#Const USE_HALFTONE = False
'    Dim LnOrigW As Long
'    Dim LnOrigH As Long
    Dim LrBitmap As BITMAP
    #If USE_HALFTONE Then
    Dim LnOldStrMode As Long
    #End If
    'Validate that a bitmap was passed in
    If oPicture.Type = vbPicTypeBitmap Then
        With oPicture
            ' get bitmap info
            GetObject .Handle, Len(LrBitmap), LrBitmap   'dest
            ' The following line was supposed to work, but...
            'HiMetricToPixels .Width, .Height, LnOrigW, LnOrigH
            'Create a DC to select bitmap into
            #If USE_DISPLAY Then
            Dim LhDC As Long
            LhDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
            LhDCTemp = CreateCompatibleDC(LhDC)
            DeleteDC LhDC
            #Else
            LhDCTemp = CreateCompatibleDC(m_hWorkDC)
            #End If
            #If USE_PALETTE Then
            LhOldPalette = SelectPalette(LhDCTemp, .hPal, True)
            RealizePalette LhDCTemp
            #End If
            'Select bitmap into DC
            LhOldBitmap = SelectObject(LhDCTemp, .Handle)
        End With
        #If USE_BITBLT Then
        If ((lWidth = LrBitmap.bmWidth) And (lHeight = LrBitmap.bmHeight)) Then
            #If USE_HALFTONE Then
            LnOldStrMode = SetStretchBltMode(m_hWorkDC, HALFTONE)
            #End If
            BitBlt m_hWorkDC, lLeft, lTop, lWidth, lHeight, LhDCTemp, 0, 0, SRCCOPY
            #If USE_HALFTONE Then
            LnOldStrMode = SetStretchBltMode(m_hWorkDC, LnOldStrMode)
            #End If
        Else
        #Else
        If True Then
        #End If
            With LrReal
                If bKeepRatio Then
                    Dim LnRatio As Single
                    Dim LnDrwRatio As Single
                    Dim LnImgRatio As Single
                    
                    With LrBitmap
                        LnDrwRatio = (lHeight / lWidth)
                        LnImgRatio = (.bmHeight / .bmWidth)
                        If (LnDrwRatio > LnImgRatio) Then
                            LnRatio = (lWidth / .bmWidth)
                        Else
                            LnRatio = (lHeight / .bmHeight)
                        End If
                    End With
                    ' Used for width & height
                    .Right = (LrBitmap.bmWidth * LnRatio)
                    .Bottom = (LrBitmap.bmHeight * LnRatio)
                    ' Left & Top positions
                    If bHCenter Then
                        .Left = (lLeft + ((lWidth - .Right) / 2))
                    Else
                        .Left = lLeft
                    End If
                    If bVCenter Then
                        .Top = (lTop + ((lHeight - .Bottom) / 2))
                    Else
                        .Top = lTop
                    End If
                Else
                    .Left = lLeft
                    .Top = lTop
                    ' Used for width & height
                    .Right = lWidth
                    .Bottom = lHeight
                End If
                'Copy to destination DC
                #If USE_BKMODE Then
                Dim LnOldBkColor As Long
                LnOldBkColor = SetBkColor(m_hWorkDC, 0&)
                #End If
                #If USE_HALFTONE Then
                LnOldStrMode = SetStretchBltMode(m_hWorkDC, HALFTONE)
                #End If
                Call SetBrushOrgEx(m_hWorkDC, 0, 0, LrPoint)
                Call StretchBlt(m_hWorkDC, .Left, .Top, .Right, .Bottom, LhDCTemp, 0, 0, LrBitmap.bmWidth, LrBitmap.bmHeight, SRCCOPY)
                #If USE_BKMODE Then
                LnOldBkColor = SetBkColor(m_hWorkDC, LnOldBkColor)
                #End If
                #If USE_HALFTONE Then
                LnOldStrMode = SetStretchBltMode(m_hWorkDC, LnOldStrMode)
                #End If
            End With
        End If
        'Cleanup
        SelectObject LhDCTemp, LhOldBitmap
        #If USE_PALETTE Then
        SelectPalette LhDCTemp, LhOldPalette, True
        RealizePalette LhDCTemp
        #End If
        DeleteDC LhDCTemp
    Else
        Me.RaiseErr ecInvalidObjType, "DrawPicture", "Invalid picture format."
    End If
    Exit Sub
ERR_H:
    Me.RaiseErr Err.Number, "DrawPicture"
End Sub

Friend Function HiMetricToPixels(ByVal lHiMX As Long, _
    ByVal lHiMY As Long, PixelsX As Long, PixelsY As Long) As Boolean
'////////////////////
'/// DIDN'T WORK...   =[
'////////////////////
'    Dim LrMap As POINTAPI ' Stores a point (x,y) to help convert from HIMETRIC to pixels.
'    Dim LnOldMapMode As Integer ' Stores the old map mode.
'    Dim LnRetCD As Long ' Holds API return values.
'    Dim lHDC As Long
'
'    lHDC = GetDC(0&)
'    ' Set the mapmode of the hdc to HIMETRIC and calculate the size in pixels
'    LnOldMapMode = SetMapMode(lHDC, MM_HIMETRIC)
'    ' Save hi-metric values
'    LrMap.X = lHiMX
'    LrMap.Y = lHiMY
'    ' LPtoDP will convert the width and height to pixels.
'    LnRetCD = LPtoDP(lHDC, LrMap, 1)
'    If (LnRetCD = 0) Then
'       ' Restore LnOldMapMode and exit
'        GoTo Finalize
'    End If
'    ' Return the result
'    PixelsX = Abs(LrMap.X)
'    PixelsY = Abs(LrMap.Y)
'Finalize:
'    ' Restore the mapping mode as it was
'    ' & release the device context
'    Select Case LnOldMapMode
'    Case 0, -1 'invalid_handle_value
'    Case Else
'        LnRetCD = SetMapMode(lHDC, LnOldMapMode)
'        ReleaseDC 0&, lHDC
'    End Select
End Function

Friend Sub Cls()
    Dim r As RECT
    Dim lOldBackColor As Long
    Dim lBackColor As Long
'   ***** puts the dimensions into a rect structure
    r.Left = 0
    r.Top = 0
    r.Bottom = m_lHeight
    r.Right = m_lWidth
'   *****
    lOldBackColor = SetBkColor(m_hWorkDC, BackColor) ' set the back color of the dc
    ' fills the rect with the back color. Using ExtTextOut (without text) is quicker than FillRect
    ExtTextOut m_hWorkDC, 0, 0, ETO_CLIPPED Or ETO_OPAQUE, r, "", 0, 0
    SetBkColor m_hWorkDC, lOldBackColor ' reset the old backcolor
End Sub


Public Property Get hDC() As Long
   hDC = m_hWorkDC
End Property

Public Sub SetFont(sFaceName As String, nSize As Single, Optional bBold As Boolean, _
    Optional bItalic As Boolean, Optional bStrikeOut As Boolean, Optional bUnderline As Boolean, _
    Optional iRotation As Integer = 0)
    Dim LrFont As LOGFONT
    Dim LaTempArray() As Byte ' byte array to hold the fontname
    Dim LnIdx As Integer
    Dim LhFont As Long
    
    ' Restores original font
    If (Not (m_lOldFont = 0)) Then
        SelectObject m_hWorkDC, m_lOldFont
        m_lOldFont = 0
    End If
    With LrFont
        ' All but two properties are very straight-forward,
        ' even with rotation, and map directly.
        .lfHeight = -(nSize * GetDeviceCaps(m_hWorkDC, LOGPIXELSY)) / TWIPS_PER_POINT
        .lfWidth = 0
        .lfEscapement = (iRotation * 10)
        .lfOrientation = .lfEscapement
        If bBold Then ' if true set the weight to the appropriate value
            .lfWeight = FW_BOLD
        Else
            .lfWeight = FW_NORMAL
        End If
        .lfItalic = Abs(bItalic)
        .lfUnderline = Abs(bUnderline)
        .lfStrikeOut = Abs(bStrikeOut)
        .lfClipPrecision = CLIP_DEFAULT_PRECIS
        .lfQuality = PROOF_QUALITY
        .lfPitchAndFamily = DEFAULT_PITCH Or FF_DONTCARE
        ' OEM fonts can't rotate, and we must force
        ' substitution with something ANSI.
        .lfCharSet = DEFAULT_CHARSET
        If .lfCharSet = OEM_CHARSET Then
            If (iRotation <> 0) Then
                .lfCharSet = ANSI_CHARSET
            End If
        End If
        ' Only TrueType fonts can rotate, so we must
        ' specify TT-only if angle is not zero.
        If (iRotation <> 0) Then
            .lfOutPrecision = OUT_TT_ONLY_PRECIS
        Else
            .lfOutPrecision = OUT_DEFAULT_PRECIS
        End If
        ' converts the font name from unicode
        LaTempArray = StrConv(sFaceName & vbNullChar, vbFromUnicode)
    '   ***** puts the font name into the byte array for face name
        For LnIdx = 0 To UBound(LaTempArray)
            .lfFaceName(LnIdx) = LaTempArray(LnIdx)
        Next LnIdx
    End With
    LhFont = CreateFontIndirect(LrFont)
    m_lOldFont = SelectObject(m_hWorkDC, LhFont)
End Sub

Public Property Let Width(ByVal lWidth As Long)
    If (lWidth > m_lWidth) Then
        Create lWidth, m_lHeight
    End If
End Property

Public Property Get Width() As Long
    Width = m_lWidth
End Property


Public Property Let Height(ByVal lHeight As Long)
    If (lHeight > m_lHeight) Then
        Create m_lWidth, lHeight
    End If
End Property

Public Property Get Height() As Long
    Height = m_lHeight
End Property


Friend Sub Create(ByVal lWidth As Long, ByVal lHeight As Long)
' Creates the memory DC
    Dim LhDC As Long
    
    prvDestroy
    LhDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    m_hWorkDC = CreateCompatibleDC(LhDC)
    m_hBmp = CreateCompatibleBitmap(LhDC, lWidth, lHeight)
    m_hBmpOld = SelectObject(m_hWorkDC, m_hBmp)
    If (m_hBmpOld = 0) Then
        prvDestroy
    Else
        m_lWidth = lWidth
        m_lHeight = lHeight
        Cls
    End If
    DeleteDC LhDC
End Sub

Private Sub prvDestroy()
    ' Restores original font
    If (Not (m_lOldFont = 0)) Then
        SelectObject m_hWorkDC, m_lOldFont
        m_lOldFont = 0
    End If
    ' Restores original bitmap
    If (Not (m_hBmpOld = 0)) Then
        SelectObject m_hWorkDC, m_hBmpOld
        m_hBmpOld = 0
    End If
    ' Destroy created bitmap
    If (Not (m_hBmp = 0)) Then
        DeleteObject m_hBmp
        m_hBmp = 0
    End If
    ' Resets Width & Height
    m_lWidth = 0
    m_lHeight = 0
    ' Destroy created memory DC
    If (Not (m_hWorkDC = 0)) Then
        DeleteDC m_hWorkDC
        m_hWorkDC = 0
    End If
End Sub

Private Sub Class_Initialize()
    BackColor = vbWhite
    m_iScaleMode = vbInches
'    m_iMFactor=
End Sub

Public Property Let ScaleMode(iUnits As ScaleModeConstants)
'    On Error GoTo ERR_H
'    m_iScaleMode = iUnits
'    Select Case iUnits
'        Case vbPixels
'            m_iMFactor = TWIPS_PER_PIXEL
'        Case rdCentimeters
'            m_iMFactor = TWIPS_PER_CENTIMETER
'        Case rdPoints
'            m_iMFactor = TWIPS_PER_POINT
'        Case rdTwips
'            m_iMFactor = TWIPS_PER_TWIP
'        Case Else ' rdInches
'            m_iMFactor = TWIPS_PER_INCHE
'    End Select
'    m_iMFactor = (m_iMFactor / TWIPS_PER_PIXEL)
'    Exit Property
'ERR_H:
'    Me.RaiseErr Err.Number, "ScaleMode[Let]", Err.Description
End Property


Private Sub Class_Terminate()
   prvDestroy
End Sub

