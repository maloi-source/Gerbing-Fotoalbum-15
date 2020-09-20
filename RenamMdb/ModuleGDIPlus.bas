Attribute VB_Name = "ModuleGDIPlus"
Option Explicit

    ' ===============================================================
    ' Ben�tigte GDIPlus-Deklarationen zum Ermitteln Bildbreite Bildh�he
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
    
    Public Type PICTDESC
        cbSizeOfStruct As Long
        picType As Long
        hgdiObj As Long
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
        
    Public Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal image As Long, graphics As Long) As Status
    Public Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, ByVal hpal As Long, Bitmap As Long) As Status
    Public Declare Function GdipGetImageThumbnail Lib "gdiplus" (ByVal image As Long, ByVal thumbWidth As Long, ByVal thumbHeight As Long, thumbImage As Long, _
                        Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As Status
    Public Declare Function GdipDrawImageRectI Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long) As Status
    Public Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" _
                        (ByVal Bitmap As Long, ByRef hbmReturn As Long, _
                        ByVal background As Long) As Status
                        
    Private Declare Sub OleCreatePictureIndirect Lib "oleaut32.dll" _
                        (lpPictDesc As PICTDESC, riid As IID, ByVal fOwn As Boolean, _
                        lplpvObj As Object)



    
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
    
    Public m_lngInstance As Long
    Public m_lngGraphics As Long

Public Function LoadPicBox(ByVal Dateiname As String) As Long
            
    'GdipGetImageDimension ist eine schnelle Funktion zum Ermitteln von Bildbreite und Bildh�he
    'm�glich mit BMP, DIB, JPG, GIF, PNG, TIF
    'nicht m�glich mit CUR, PSD
    'Falsches Ergebnis bei EMF, WMF

    Dim retcode As Long              ' Funktions-R�ckgaben
    Dim Bitmap As Long
    Dim GDIP_Connection As Long      ' Verbindung zu GDIPlus
    Dim GDIP_Startup As GDIPlusStartupInput
    Dim MyDateiname As String
      
    On Error GoTo exitfunction
    Err.Clear
    
    'Dateiname = gstrFRODN
    MyDateiname = Dateiname
             
    gsngPicWidth = 0                'wenn diese Werte = 0 bleiben, lag ein Fehler vor
    gsngPicHeight = 0
    GDIP_Startup.GdiplusVersion = 1
    retcode = GdiplusStartup(GDIP_Connection, GDIP_Startup, ByVal 0&)
    If retcode <> 0 Then
       GoTo exitfunction
    End If
      
    ' Tr�gt das Bild aus der Datei in die Bitmap ein
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
      ' Bitmap l�schen
      GdipDisposeImage Bitmap
    End If
      
    If GDIP_Connection <> 0 Then
      ' GDIPlus-DLL freigeben
      GdiplusShutdown GDIP_Connection
    End If
    LoadPicBox = retcode
End Function

'------------------------------------------------------
' Funktion     : CreateThumbnailFromFile
' Beschreibung : L�dt ein Bilddatei per GDI+
' �bergabewert : FileName = Pfad\Dateiname der Bilddatei
'                Percent = Gr��e in Prozent (100% = 1:1)
' R�ckgabewert : StdPicture Objekt
'------------------------------------------------------
Public Function CreateThumbnailFromFile(ByVal FileName As String, _
ByVal Percent As Long) As StdPicture
    
    Dim retStatus As Status
    Dim lBitmap As Long
    Dim lThumb As Long
    Dim hBitmap As Long
    Dim ImageWidth As Single
    Dim ImageHeight As Single
    Dim IW As Long
    Dim IH As Long
    
    ' �ffnet die Bilddatei in lBitmap
    retStatus = Execute(GdipLoadImageFromFile(StrPtr(FileName), _
        lBitmap))
    
    If retStatus = OK Then
        
        ' Ermitteln der ImageDimensionen
        Call Execute(GdipGetImageDimension(lBitmap, ImageWidth, _
            ImageHeight))
            
        IW = (ImageWidth * Percent) \ 100
        IH = (ImageHeight * Percent) \ 100
        
        ' Thumbnail erzeugen
        retStatus = Execute(GdipGetImageThumbnail(lBitmap, IW, IH, _
            lThumb, 0, 0))
        
        If retStatus = Status.OK Then
            
            ' Erzeugen der GDI Bitmap von der Thumbnail Bitmap
            retStatus = Execute(GdipCreateHBITMAPFromBitmap(lThumb, _
                hBitmap, 0))
            
            If retStatus = Status.OK Then
                
                ' Erzeugen des StdPicture Objekts von hBitmap
                Set CreateThumbnailFromFile = _
                    HandleToPicture(hBitmap, vbPicTypeBitmap)
            End If
            
            ' L�sche lThumb
            Call Execute(GdipDisposeImage(lThumb))
        End If
        
        ' L�sche lBitmap
        Call Execute(GdipDisposeImage(lBitmap))
    End If
End Function

'------------------------------------------------------
' Funktion     : HandleToPicture
' Beschreibung : Umwandeln einer GDI+ Bitmap Handle in ein
'                StdPicture Objekt
' �bergabewert : hGDIHandle = GDI+ Bitmap Handle
'                ObjectType = Bitmaptyp
' R�ckgabewert : StdPicture Objekt
'------------------------------------------------------
Public Function HandleToPicture(ByVal hGDIHandle As Long, _
    ByVal ObjectType As PictureTypeConstants, _
    Optional ByVal hpal As Long = 0) As StdPicture
    
    Dim tPictDesc As PICTDESC
    Dim IID_IPicture As IID
    Dim oPicture As IPicture
    
    ' Initialisiert die PICTDESC Structur
    With tPictDesc
        .cbSizeOfStruct = Len(tPictDesc)
        .picType = ObjectType
        .hgdiObj = hGDIHandle
        .hPalOrXYExt = hpal
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
    
    ' R�ckgabe des Pictureobjekts
    Set HandleToPicture = oPicture
End Function

'------------------------------------------------------
' Funktion     : Execute
' Beschreibung : Gibt im Fehlerfall die entsprechende GDI+ Fehlermeldung aus
' �bergabewert : GDI+ Status
' R�ckgabewert : GDI+ Status
'------------------------------------------------------
Private Function Execute(ByVal lReturn As Status) As Status
    Dim lCurErr As Status
    If lReturn = Status.OK Then
        lCurErr = Status.OK
    Else
        lCurErr = lReturn
        Call MsgBox(GdiErrorString(lReturn) & " GDI+ Error:" & lReturn, _
                     vbOKOnly, "GDI Error")
    End If
    Execute = lCurErr
End Function

'------------------------------------------------------
' Funktion     : GdiErrorString
' Beschreibung : Umwandlung der GDI+ Statuscodes in Stringcodes
' �bergabewert : GDI+ Status
' R�ckgabewert : Fehlercode als String
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
' Beschreibung : L�dt ein Bilddatei per GDI+
' �bergabewert : FileName = Pfad\Dateiname der Bilddatei
'                Percent = Gr��e in Prozent (100% = 1:1)
' R�ckgabewert : StdPicture Objekt
'------------------------------------------------------
Private Function DrawImageFromFile(ByVal FileName As String, _
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
        
        ' �ffnet die Bilddatei in lBitmap
        retStatus = Execute(GdipLoadImageFromFile(StrPtr(FileName), _
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
            
            ' L�sche lBitmap
            Call Execute(GdipDisposeImage(lBitmap))
            
        End If
        ' L�sche das Grafikobjekt
        Call Execute(GdipDeleteGraphics(lngGraphics))
    End If
End Function

Public Sub ThumbnailAnzeigen(EchterStandort As String, ByRef PicBox As VB.PictureBox)
    Dim retcode As Long
    Dim prozentH As Long
    Dim prozentW As Long
    Dim prozent As Long
    Dim udtData As GDIPlusStartupInput

    PicBox.AutoSize = False
    PicBox.AutoRedraw = True
    EchterStandort = Replace(EchterStandort, "+:\", AppPath & "\")
    retcode = LoadPicBox(EchterStandort)
    If gsngPicHeight = 0 Then
        Exit Sub
    End If
    
    udtData.GdiplusVersion = 1
    If GdiplusStartup(m_lngInstance, udtData, 0) Then
        MsgBox "GDI+ could not be initialized", vbCritical
        Exit Sub
    End If
    prozentH = (PicBox.Height / Screen.TwipsPerPixelY) / gsngPicHeight * 100            'prozentH ausrechnen
    prozentW = (PicBox.Width / Screen.TwipsPerPixelX) / gsngPicWidth * 100              'prozentW ausrechnen
    If prozentH < prozentW Then
        prozent = prozentH
    Else
        prozent = prozentW
    End If
    PicBox.Picture = LoadPicture("")
    Call DrawImageFromFile(EchterStandort, PicBox.hdc, prozent)
    PicBox.Picture = PicBox.image
    GdiplusShutdown m_lngInstance
    Exit Sub
End Sub

