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
      ByVal FileName As Long, _
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
                
    'Graphics:
    Public Declare Function GdipCreateFromHDC Lib "gdiplus.dll" ( _
        ByVal hdc As Long, ByRef graphics As Long _
        ) As Status
        
    Public Declare Function GdipDeleteGraphics Lib "gdiplus.dll" ( _
        ByVal graphics As Long _
        ) As Status
        
    Public Declare Function GdipGraphicsClear Lib "gdiplus.dll" ( _
        ByVal graphics As Long, ByVal color As Long _
        ) As Status
        
    Public Declare Function GdipDrawImage Lib "gdiplus.dll" ( _
        ByVal graphics As Long, ByVal image As Long, _
        ByVal x As Single, ByVal y As Single _
        ) As Status
                                
    Public Declare Function GdipDrawImageRect Lib "gdiplus" _
        (ByVal graphics As Long, ByVal image As Long, _
        ByVal x As Single, ByVal y As Single, ByVal Width As Single, _
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
    
    Public Enum Status
        Ok = 0
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
    
    Public m_lngInstance As Long
    Public m_lngGraphics As Long

Public Function LoadPicBox(ByVal Dateiname As String) As Long
            
    'GdipGetImageDimension ist eine schnelle Funktion zum Ermitteln von Bildbreite und Bildhöhe
    'möglich mit BMP, DIB, JPG, GIF, PNG, TIF
    'nicht möglich mit CUR, PSD
    'Falsches Ergebnis bei EMF, WMF

    Dim retcode As Long              ' Funktions-Rückgaben
    Dim Bitmap As Long
    Dim GDIP_Connection As Long      ' Verbindung zu GDIPlus
    Dim GDIP_Startup As GDIPlusStartupInput
    Dim MyDateiname As String
      
    On Error GoTo exitfunction
    Err.Clear
             
    MyDateiname = Dateiname
             
    gsngPicWidth = 0                'wenn diese Werte = 0 bleiben, lag ein Fehler vor
    gsngPicHeight = 0
    GDIP_Startup.GdiplusVersion = 1
    retcode = GdiplusStartup(GDIP_Connection, GDIP_Startup, ByVal 0&)
    If retcode <> 0 Then
       GoTo exitfunction
    End If
      
    ' Trägt das Bild aus der Datei in die Bitmap ein
    retcode = GdipLoadImageFromFile(StrPtr(MyDateiname), Bitmap)
    If retcode <> 0 Then
       GoTo exitfunction
    End If
          
    ' Abfrage der Abmessungen der Bitmap
    retcode = GdipGetImageDimension(Bitmap, gsngPicWidth, gsngPicHeight)
    If retcode <> 0 Then
      GoTo exitfunction
    End If
exitfunction:
    'MsgBox "retcode=" & retcode
    ' Ressourcen und GDIPLus freigeben
    If Bitmap <> 0 Then
      ' Bitmap löschen
      GdipDisposeImage Bitmap
    End If
      
    If GDIP_Connection <> 0 Then
      ' GDIPlus-DLL freigeben
      GdiplusShutdown GDIP_Connection
    End If
    LoadPicBox = retcode
End Function
