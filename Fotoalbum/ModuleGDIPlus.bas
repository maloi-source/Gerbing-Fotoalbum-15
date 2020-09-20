Attribute VB_Name = "ModuleGDIPlus"
Option Explicit

    ' ===============================================================
    ' Benötigte GDIPlus-Deklarationen zum Ermitteln Bildbreite Bildhöhe
    ' ===============================================================
    ' Verbindung herstellen
    Public Declare Function GdiplusStartup Lib "gdiplus" ( _
      ByRef GDIP_Connection As Long, _
      ByRef udtInput As GDIPlusStartupInput, _
      Optional ByRef udtOutput As Any) As Long
    
    ' Bild aus Datei laden (in Bitmap)
    Public Declare Function GdipLoadImageFromFile Lib "gdiplus" ( _
      ByVal Filename As Long, _
      ByRef image As Long) As Long
    
    ' Abmessungen des Bildes ermitteln
    Public Declare Function GdipGetImageDimension Lib "gdiplus" ( _
      ByVal image As Long, _
      ByRef Width As Single, _
      ByRef Height As Single) As Long
    
    ' Bild-Ressource freigeben
    Public Declare Function GdipDisposeImage Lib "gdiplus" ( _
      ByVal image As Long) As Long
        
    ' GDIPLus freigeben
    Public Declare Function GdiplusShutdown Lib "gdiplus" ( _
      ByVal token As Long) As Long
      
    Public gsngPicWidth As Single           ' Bildabmessungen
    Public gsngPicHeight As Single


    Public Const QualityModeInvalid As Long = -1&
    Public Const QualityModeDefault As Long = 0&
    Public Const QualityModeLow As Long = 1&
    Public Const QualityModeHigh As Long = 2&
    
    Public Enum InterpolationMode
        InterpolationModeInvalid = QualityModeInvalid
        InterpolationModeDefault = QualityModeDefault
        InterpolationModeLowQuality = QualityModeLow
        InterpolationModeHighQuality = QualityModeHigh
        InterpolationModeBilinear = QualityModeHigh + 1
        InterpolationModeBicubic = QualityModeHigh + 2
        InterpolationModeNearestNeighbor = QualityModeHigh + 3
        InterpolationModeHighQualityBilinear = QualityModeHigh + 4
        InterpolationModeHighQualityBicubic = QualityModeHigh + 5
    End Enum
    
    Public Enum SmoothingMode
        SmoothingModeInvalid = QualityModeInvalid
        SmoothingModeDefault = QualityModeDefault
        SmoothingModeHighSpeed = QualityModeLow
        SmoothingModeHighQuality = QualityModeHigh
        SmoothingModeNone = QualityModeHigh + 1
        SmoothingModeAntiAlias8x4 = QualityModeHigh + 2
        SmoothingModeAntiAlias = SmoothingModeAntiAlias8x4
        'SmoothingModeAntiAlias8x8
    End Enum
    
    Public Enum PixelOffsetMode
        PixelOffsetModeInvalid = QualityModeInvalid
        PixelOffsetModeDefault = QualityModeDefault
        PixelOffsetModeHighSpeed = QualityModeLow
        PixelOffsetModeHighQuality = QualityModeHigh
        PixelOffsetModeNone = QualityModeHigh + 1
        PixelOffsetModeHalf = QualityModeHigh + 2
    End Enum
    
    Public Enum CompositingQualityMode
        CompositingQualityInvalid = QualityModeInvalid
        CompositingQualityDefault = QualityModeDefault
        CompositingQualityHighSpeed = QualityModeLow
        CompositingQualityHighQuality = QualityModeHigh
        CompositingQualityGammaCorrected = QualityModeHigh + 1
        CompositingQualityAssumeLinear = QualityModeHigh + 2
    End Enum
    
    Public Enum CompositingModeMode
        CompositingModeSourceOver = 0
        CompositingModeSourceCopy = 1
    End Enum
    
    Public Type PICTDESC
        cbSizeOfStruct As Long
        picType As Long
        hgdiobj As Long
        hPalOrXYExt As Long
    End Type
    
    Private Type IID
        Data1 As Long
        Data2 As Integer
        Data3 As Integer
        Data4(0 To 7)  As Byte
    End Type


                
    'Graphics:
    Public Declare Function GdipCreateFromHDC Lib "gdiplus.dll" ( _
        ByVal hDC As Long, ByRef graphics As Long _
        ) As Status
        
    Public Declare Function GdipDeleteGraphics Lib "gdiplus.dll" ( _
        ByVal graphics As Long _
        ) As Status
        
    Public Declare Function GdipGraphicsClear Lib "gdiplus.dll" ( _
        ByVal graphics As Long, ByVal Color As Long _
        ) As Status
        
    Public Declare Function GdipDrawImage Lib "gdiplus.dll" ( _
        ByVal graphics As Long, ByVal image As Long, _
        ByVal X As Single, ByVal Y As Single _
        ) As Status
                                
    Public Declare Function GdipDrawImageRect Lib "gdiplus" _
        (ByVal graphics As Long, ByVal image As Long, _
        ByVal X As Single, ByVal Y As Single, ByVal Width As Single, _
        ByVal Height As Single) As Status
        
    Public Declare Function GdipSetSmoothingMode Lib "gdiplus" _
        (ByVal graphics As Long, ByVal SmoothingMode As _
        SmoothingMode) As Status
        
    Public Declare Function GdipSetInterpolationMode Lib "gdiplus" _
        (ByVal graphics As Long, ByVal InterpolationMode As _
        InterpolationMode) As Status
        
    Public Declare Function GdipSetPixelOffsetMode Lib "gdiplus" _
        (ByVal graphics As Long, ByVal PixelOffsetMode As _
        PixelOffsetMode) As Status
        
    Public Declare Function GdipSetCompositingQuality Lib "gdiplus" _
        (ByVal graphics As Long, ByVal CompositingQuality As _
        CompositingQualityMode) As Status
        
    Public Declare Function GdipSetCompositingMode Lib "gdiplus" _
        (ByVal graphics As Long, ByVal CompositingMode As _
        CompositingModeMode) As Status
        
    Public Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal image As Long, graphics As Long) As Status
    Public Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, ByVal hPal As Long, BITMAP As Long) As Status
    Public Declare Function GdipGetImageThumbnail Lib "gdiplus" (ByVal image As Long, ByVal thumbWidth As Long, ByVal thumbHeight As Long, thumbImage As Long, _
                        Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As Status
    Public Declare Function GdipDrawImageRectI Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Status
    Public Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" _
                        (ByVal BITMAP As Long, ByRef hbmReturn As Long, _
                        ByVal background As Long) As Status
                        
    Private Declare Sub OleCreatePictureIndirect Lib "oleaut32.dll" _
                        (lpPictDesc As PICTDESC, riid As IID, ByVal fOwn As Boolean, _
                        lplpvObj As Object)
                        
    Public Declare Function GdipCreateHatchBrush Lib "gdiplus" (ByVal style As HatchStyle, ByVal forecolr As Long, ByVal backcolr As Long, brush As Long) As Status
    Public Declare Function GdipFillRectangle Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Status
    
    Public Enum Status
        OK = 0
        GenericError = 1
        InvalidParameter = 2
        OutOfMemory = 3
        ObjectBusy = 4
        InsufficientBuffer = 5
        NotImplemented = 6
        Win32Error = 7
        WrongState = 8
        Aborted = 9
        FileNotFound = 10
        ValueOverflow = 11
        AccessDenied = 12
        UnknownImageFormat = 13
        FontFamilyNotFound = 14
        FontStyleNotFound = 15
        NotTrueTypeFont = 16
        UnsupportedGdiplusVersion = 17
        GdiplusNotInitialized = 18
        PropertyNotFound = 19
        PropertyNotSupported = 20
        ProfileNotFound = 21
    End Enum
    
    Public Type GDIPlusStartupInput
        GdiplusVersion As Long
        DebugEventCallback As Long
        SuppressBackgroundThread As Long
        SuppressExternalCodecs As Long
    End Type
    
    ' Various Hatch Styles
    Public Enum HatchStyle
       HatchStyleHorizontal                   ' 0
       HatchStyleVertical                     ' 1
       HatchStyleForwardDiagonal              ' 2
       HatchStyleBackwardDiagonal             ' 3
       HatchStyleCross                        ' 4
       HatchStyleDiagonalCross                ' 5
       HatchStyle05Percent                    ' 6
       HatchStyle10Percent                    ' 7
       HatchStyle20Percent                    ' 8
       HatchStyle25Percent                    ' 9
       HatchStyle30Percent                    ' 10
       HatchStyle40Percent                    ' 11
       HatchStyle50Percent                    ' 12
       HatchStyle60Percent                    ' 13
       HatchStyle70Percent                    ' 14
       HatchStyle75Percent                    ' 15
       HatchStyle80Percent                    ' 16
       HatchStyle90Percent                    ' 17
       HatchStyleLightDownwardDiagonal        ' 18
       HatchStyleLightUpwardDiagonal          ' 19
       HatchStyleDarkDownwardDiagonal         ' 20
       HatchStyleDarkUpwardDiagonal           ' 21
       HatchStyleWideDownwardDiagonal         ' 22
       HatchStyleWideUpwardDiagonal           ' 23
       HatchStyleLightVertical                ' 24
       HatchStyleLightHorizontal              ' 25
       HatchStyleNarrowVertical               ' 26
       HatchStyleNarrowHorizontal             ' 27
       HatchStyleDarkVertical                 ' 28
       HatchStyleDarkHorizontal               ' 29
       HatchStyleDashedDownwardDiagonal       ' 30
       HatchStyleDashedUpwardDiagonal         ' 31
       HatchStyleDashedHorizontal             ' 32
       HatchStyleDashedVertical               ' 33
       HatchStyleSmallConfetti                ' 34
       HatchStyleLargeConfetti                ' 35
       HatchStyleZigZag                       ' 36
       HatchStyleWave                         ' 37
       HatchStyleDiagonalBrick                ' 38
       HatchStyleHorizontalBrick              ' 39
       HatchStyleWeave                        ' 40
       HatchStylePlaid                        ' 41
       HatchStyleDivot                        ' 42
       HatchStyleDottedGrid                   ' 43
       HatchStyleDottedDiamond                ' 44
       HatchStyleShingle                      ' 45
       HatchStyleTrellis                      ' 46
       HatchStyleSphere                       ' 47
       HatchStyleSmallGrid                    ' 48
       HatchStyleSmallCheckerBoard            ' 49
       HatchStyleLargeCheckerBoard            ' 50
       HatchStyleOutlinedDiamond              ' 51
       HatchStyleSolidDiamond                 ' 52
    
       HatchStyleTotal
       HatchStyleLargeGrid = HatchStyleCross  ' 4
    
       HatchStyleMin = HatchStyleHorizontal
       HatchStyleMax = HatchStyleTotal - 1
    End Enum

    ' Common color constants
    ' NOTE: Oringinal enum was unnamed
    Public Enum Colors
       AliceBlue = &HFFF0F8FF
       AntiqueWhite = &HFFFAEBD7
       Aqua = &HFF00FFFF
       Aquamarine = &HFF7FFFD4
       Azure = &HFFF0FFFF
       Beige = &HFFF5F5DC
       Bisque = &HFFFFE4C4
       Black = &HFF000000
       BlanchedAlmond = &HFFFFEBCD
       Blue = &HFF0000FF
       BlueViolet = &HFF8A2BE2
       Brown = &HFFA52A2A
       BurlyWood = &HFFDEB887
       CadetBlue = &HFF5F9EA0
       Chartreuse = &HFF7FFF00
       Chocolate = &HFFD2691E
       Coral = &HFFFF7F50
       CornflowerBlue = &HFF6495ED
       Cornsilk = &HFFFFF8DC
       Crimson = &HFFDC143C
       Cyan = &HFF00FFFF
       DarkBlue = &HFF00008B
       DarkCyan = &HFF008B8B
       DarkGoldenrod = &HFFB8860B
       DarkGray = &HFFA9A9A9
       DarkGreen = &HFF006400
       DarkKhaki = &HFFBDB76B
       DarkMagenta = &HFF8B008B
       DarkOliveGreen = &HFF556B2F
       DarkOrange = &HFFFF8C00
       DarkOrchid = &HFF9932CC
       DarkRed = &HFF8B0000
       DarkSalmon = &HFFE9967A
       DarkSeaGreen = &HFF8FBC8B
       DarkSlateBlue = &HFF483D8B
       DarkSlateGray = &HFF2F4F4F
       DarkTurquoise = &HFF00CED1
       DarkViolet = &HFF9400D3
       DeepPink = &HFFFF1493
       DeepSkyBlue = &HFF00BFFF
       DimGray = &HFF696969
       DodgerBlue = &HFF1E90FF
       Firebrick = &HFFB22222
       FloralWhite = &HFFFFFAF0
       ForestGreen = &HFF228B22
       Fuchsia = &HFFFF00FF
       Gainsboro = &HFFDCDCDC
       GhostWhite = &HFFF8F8FF
       Gold = &HFFFFD700
       Goldenrod = &HFFDAA520
       Gray = &HFF808080
       Green = &HFF008000
       GreenYellow = &HFFADFF2F
       Honeydew = &HFFF0FFF0
       HotPink = &HFFFF69B4
       IndianRed = &HFFCD5C5C
       Indigo = &HFF4B0082
       Ivory = &HFFFFFFF0
       Khaki = &HFFF0E68C
       Lavender = &HFFE6E6FA
       LavenderBlush = &HFFFFF0F5
       LawnGreen = &HFF7CFC00
       LemonChiffon = &HFFFFFACD
       LightBlue = &HFFADD8E6
       LightCoral = &HFFF08080
       LightCyan = &HFFE0FFFF
       LightGoldenrodYellow = &HFFFAFAD2
       LightGray = &HFFD3D3D3
       LightGreen = &HFF90EE90
       LightPink = &HFFFFB6C1
       LightSalmon = &HFFFFA07A
       LightSeaGreen = &HFF20B2AA
       LightSkyBlue = &HFF87CEFA
       LightSlateGray = &HFF778899
       LightSteelBlue = &HFFB0C4DE
       LightYellow = &HFFFFFFE0
       Lime = &HFF00FF00
       LimeGreen = &HFF32CD32
       Linen = &HFFFAF0E6
       Magenta = &HFFFF00FF
       Maroon = &HFF800000
       MediumAquamarine = &HFF66CDAA
       MediumBlue = &HFF0000CD
       MediumOrchid = &HFFBA55D3
       MediumPurple = &HFF9370DB
       MediumSeaGreen = &HFF3CB371
       MediumSlateBlue = &HFF7B68EE
       MediumSpringGreen = &HFF00FA9A
       MediumTurquoise = &HFF48D1CC
       MediumVioletRed = &HFFC71585
       MidnightBlue = &HFF191970
       MintCream = &HFFF5FFFA
       MistyRose = &HFFFFE4E1
       Moccasin = &HFFFFE4B5
       NavajoWhite = &HFFFFDEAD
       Navy = &HFF000080
       OldLace = &HFFFDF5E6
       Olive = &HFF808000
       OliveDrab = &HFF6B8E23
       Orange = &HFFFFA500
       OrangeRed = &HFFFF4500
       Orchid = &HFFDA70D6
       PaleGoldenrod = &HFFEEE8AA
       PaleGreen = &HFF98FB98
       PaleTurquoise = &HFFAFEEEE
       PaleVioletRed = &HFFDB7093
       PapayaWhip = &HFFFFEFD5
       PeachPuff = &HFFFFDAB9
       Peru = &HFFCD853F
       Pink = &HFFFFC0CB
       Plum = &HFFDDA0DD
       PowderBlue = &HFFB0E0E6
       Purple = &HFF800080
       Red = &HFFFF0000
       RosyBrown = &HFFBC8F8F
       RoyalBlue = &HFF4169E1
       SaddleBrown = &HFF8B4513
       Salmon = &HFFFA8072
       SandyBrown = &HFFF4A460
       SeaGreen = &HFF2E8B57
       SeaShell = &HFFFFF5EE
       Sienna = &HFFA0522D
       Silver = &HFFC0C0C0
       SkyBlue = &HFF87CEEB
       SlateBlue = &HFF6A5ACD
       SlateGray = &HFF708090
       Snow = &HFFFFFAFA
       SpringGreen = &HFF00FF7F
       SteelBlue = &HFF4682B4
       'Tan = &HFFD2B48C                            'Gerbing 05.09.2016 auskommentieren sonst Fehler in ucGMap.Public Function LatToPxlY
       Teal = &HFF008080
       Thistle = &HFFD8BFD8
       Tomato = &HFFFF6347
       Transparent = &HFFFFFF
       Turquoise = &HFF40E0D0
       Violet = &HFFEE82EE
       Wheat = &HFFF5DEB3
       white = &HFFFFFFFF
       WhiteSmoke = &HFFF5F5F5
       Yellow = &HFFFFFF00
       YellowGreen = &HFF9ACD32
    End Enum
    
    Public m_lngInstance As Long
    Public m_lngGraphics As Long

Public Function LoadPicBox(ByVal Dateiname As String) As Long
            
    'GdipGetImageDimension ist eine schnelle Funktion zum Ermitteln von Bildbreite und Bildhöhe
    'möglich mit BMP, DIB, JPG, GIF, PNG, TIF
    'nicht möglich mit CUR, PSD
    'Falsches Ergebnis bei EMF, WMF

    Dim retcode As Long              ' Funktions-Rückgaben
    Dim BITMAP As Long
    Dim GDIP_Connection As Long      ' Verbindung zu GDIPlus
    Dim GDIP_Startup As GDIPlusStartupInput
    Dim MyDateiname As String
      
    On Error GoTo exitfunction
    Err.Clear
    
    'Dateiname = gstrFRODN
    MyDateiname = Dateiname
    'MessageBoxW 0, StrPtr(MyDateiname), StrPtr("GERBING Fotoalbum"), vbInformation
             
    gsngPicWidth = 0                'wenn diese Werte = 0 bleiben, lag ein Fehler vor
    gsngPicHeight = 0
    GDIP_Startup.GdiplusVersion = 1
    retcode = GdiplusStartup(GDIP_Connection, GDIP_Startup, ByVal 0&)
    If retcode <> 0 Then
       GoTo exitfunction
    End If
      
    ' Trägt das Bild aus der Datei in die Bitmap ein
    retcode = GdipLoadImageFromFile(StrPtr(MyDateiname), BITMAP)
    If retcode <> 0 Then
       GoTo exitfunction
    End If
          
    ' Abfrage der Abmessungen der Bitmap
    retcode = GdipGetImageDimension(BITMAP, gsngPicWidth, gsngPicHeight)
    If retcode <> 0 Then
      GoTo exitfunction
    End If
exitfunction:
    'MsgBox "retcode=" & retcode
    ' Ressourcen und GDIPLus freigeben
    If BITMAP <> 0 Then
      ' Bitmap löschen
      GdipDisposeImage BITMAP
    End If
      
    If GDIP_Connection <> 0 Then
      ' GDIPlus-DLL freigeben
      GdiplusShutdown GDIP_Connection
    End If
    LoadPicBox = retcode
End Function

'------------------------------------------------------
' Funktion     : CreateThumbnailFromFile
' Beschreibung : Lädt ein Bilddatei per GDI+
' Übergabewert : FileName = Pfad\Dateiname der Bilddatei
'                Percent = Größe in Prozent (100% = 1:1)
' Rückgabewert : StdPicture Objekt
'------------------------------------------------------
Public Function CreateThumbnailFromFile(ByVal Filename As String, _
ByVal Percent As Long) As StdPicture
    
    Dim retStatus As Status
    Dim lBitmap As Long
    Dim lThumb As Long
    Dim hBitmap As Long
    Dim ImageWidth As Single
    Dim ImageHeight As Single
    Dim IW As Long
    Dim IH As Long
    
    ' Öffnet die Bilddatei in lBitmap
    retStatus = Execute(GdipLoadImageFromFile(StrPtr(Filename), lBitmap))
    If retStatus = OK Then
        ' Ermitteln der ImageDimensionen
        Call Execute(GdipGetImageDimension(lBitmap, ImageWidth, ImageHeight))
        IW = (ImageWidth * Percent) \ 100
        IH = (ImageHeight * Percent) \ 100
        ' Thumbnail erzeugen
        retStatus = Execute(GdipGetImageThumbnail(lBitmap, IW, IH, _
            lThumb, 0, 0))
        If retStatus = Status.OK Then
            ' Erzeugen der GDI Bitmap von der Thumbnail Bitmap
            retStatus = Execute(GdipCreateHBITMAPFromBitmap(lThumb, hBitmap, 0))
            If retStatus = Status.OK Then
                ' Erzeugen des StdPicture Objekts von hBitmap
                Set CreateThumbnailFromFile = HandleToPicture(hBitmap, vbPicTypeBitmap)
            End If
            ' Lösche lThumb
            Call Execute(GdipDisposeImage(lThumb))
        End If
        ' Lösche lBitmap
        Call Execute(GdipDisposeImage(lBitmap))
    End If
End Function

'------------------------------------------------------
' Funktion     : HandleToPicture
' Beschreibung : Umwandeln einer GDI+ Bitmap Handle in ein
'                StdPicture Objekt
' Übergabewert : hGDIHandle = GDI+ Bitmap Handle
'                ObjectType = Bitmaptyp
' Rückgabewert : StdPicture Objekt
'------------------------------------------------------
Public Function HandleToPicture(ByVal hGDIHandle As Long, _
    ByVal ObjectType As PictureTypeConstants, _
    Optional ByVal hPal As Long = 0) As StdPicture
    
    Dim tPictDesc As PICTDESC
    Dim IID_IPicture As IID
    Dim oPicture As IPicture
    
    ' Initialisiert die PICTDESC Structur
    With tPictDesc
        .cbSizeOfStruct = Len(tPictDesc)
        .picType = ObjectType
        .hgdiobj = hGDIHandle
        .hPalOrXYExt = hPal
    End With
    
    ' Initialisiert das IPicture Interface ID
    With IID_IPicture
        .Data1 = &H7BF80981
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(3) = &HAA
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    
    ' Erzeugen des Objekts
    OleCreatePictureIndirect tPictDesc, IID_IPicture, True, oPicture
    
    ' Rückgabe des Pictureobjekts
    Set HandleToPicture = oPicture
End Function

'------------------------------------------------------
' Funktion     : Execute
' Beschreibung : Gibt im Fehlerfall die entsprechende GDI+ Fehlermeldung aus
' Übergabewert : GDI+ Status
' Rückgabewert : GDI+ Status
'------------------------------------------------------
Private Function Execute(ByVal lReturn As Status) As Status
    Dim lCurErr As Status
    If lReturn = Status.OK Then
        lCurErr = Status.OK
    Else
        lCurErr = lReturn
'        Call MsgBox(GdiErrorString(lReturn) & " GDI+ Error:" & lReturn, _
'                     vbOKOnly, "GDI Error")
    End If
    Execute = lCurErr
End Function

'------------------------------------------------------
' Funktion     : GdiErrorString
' Beschreibung : Umwandlung der GDI+ Statuscodes in Stringcodes
' Übergabewert : GDI+ Status
' Rückgabewert : Fehlercode als String
'------------------------------------------------------
Private Function GdiErrorString(ByVal lError As Status) As String
    Dim S As String
    
    Select Case lError
    Case GenericError:              S = "Generic Error."
    Case InvalidParameter:          S = "Invalid Parameter."
    Case OutOfMemory:               S = "Out Of Memory."
    Case ObjectBusy:                S = "Object Busy."
    Case InsufficientBuffer:        S = "Insufficient Buffer."
    Case NotImplemented:            S = "Not Implemented."
    Case Win32Error:                S = "Win32 Error."
    Case WrongState:                S = "Wrong State."
    Case Aborted:                   S = "Aborted."
    Case FileNotFound:              S = "File Not Found."
    Case ValueOverflow:             S = "Value Overflow."
    Case AccessDenied:              S = "Access Denied."
    Case UnknownImageFormat:        S = "Unknown Image Format."
    Case FontFamilyNotFound:        S = "FontFamily Not Found."
    Case FontStyleNotFound:         S = "FontStyle Not Found."
    Case NotTrueTypeFont:           S = "Not TrueType Font."
    Case UnsupportedGdiplusVersion: S = "Unsupported Gdiplus Version."
    Case GdiplusNotInitialized:     S = "Gdiplus Not Initialized."
    Case PropertyNotFound:          S = "Property Not Found."
    Case PropertyNotSupported:      S = "Property Not Supported."
    Case Else:                      S = "Unknown GDI+ Error."
    End Select
    
    GdiErrorString = S
End Function


'------------------------------------------------------
' Funktion     : DrawImageFromFile
' Beschreibung : Lädt ein Bilddatei per GDI+
' Übergabewert : FileName = Pfad\Dateiname der Bilddatei
'                Percent = Größe in Prozent (100% = 1:1)
' Rückgabewert : StdPicture Objekt
'------------------------------------------------------
Public Function DrawImageFromFile(ByVal Filename As String, _
    ByVal DrawHdc As Long, ByVal Percent As Long, _
    Optional ByVal Interpolation As InterpolationMode = _
    InterpolationModeDefault, Optional ByVal Smoothing As SmoothingMode _
    = SmoothingModeNone, Optional ByVal PixelOffset As PixelOffsetMode = _
    PixelOffsetModeNone, Optional ByVal CompositingQuality As _
    CompositingQualityMode = CompositingQualityDefault, Optional ByVal _
    CompositingMode As CompositingModeMode = CompositingModeSourceOver) _
    As Boolean
    
    Dim retStatus As Status
    Dim lBitmap As Long
    Dim lngGraphics As Long
    Dim ImageWidth As Single
    Dim ImageHeight As Single
    Dim IW As Single
    Dim IH As Single
    
    ' Erzeugen eines Grafikobjekts von DrawHdc -> lngGraphics
    retStatus = Execute(GdipCreateFromHDC(DrawHdc, lngGraphics))
    If retStatus = OK Then
        
        ' Setzen der Optimierungsmodis
        Call Execute(GdipSetSmoothingMode(lngGraphics, _
            Smoothing))
            
        Call Execute(GdipSetInterpolationMode(lngGraphics, _
            Interpolation))
            
        Call Execute(GdipSetPixelOffsetMode(lngGraphics, _
            PixelOffset))
        
        Call Execute(GdipSetCompositingQuality(lngGraphics, _
            CompositingQuality))
            
        Call Execute(GdipSetCompositingMode(lngGraphics, _
            CompositingMode))
        
        ' Öffnet die Bilddatei in lBitmap
        retStatus = Execute(GdipLoadImageFromFile(StrPtr(Filename), _
            lBitmap))
        
        If retStatus = OK Then
            
            ' Ermitteln der ImageDimensionen
            Call Execute(GdipGetImageDimension(lBitmap, ImageWidth, _
                ImageHeight))
                
            IW = (ImageWidth * Percent) \ 100
            IH = (ImageHeight * Percent) \ 100
            
            ' Image erzeugen
            retStatus = Execute(GdipDrawImageRect(lngGraphics, lBitmap, _
                                                  0, 0, IW, IH))
            
            ' Lösche lBitmap
            Call Execute(GdipDisposeImage(lBitmap))
            
        End If
        ' Lösche das Grafikobjekt
        Call Execute(GdipDeleteGraphics(lngGraphics))
    End If
End Function


