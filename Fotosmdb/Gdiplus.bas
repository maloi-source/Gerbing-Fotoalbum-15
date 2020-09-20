Attribute VB_Name = "Gdiplus"
'Dieser Source stammt von http://www.activevb.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.

'Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
'Ansonsten viel Spaß und Erfolg mit diesem Source!
Option Explicit

    ' ----==== GDIPlus Const ====----
    Public Const GdiPlusVersion As Long = 1
    
    ' ----==== Sonstige Types ====----
    Private Type PICTDESC
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
    
    ' ----==== GDIPlus Types ====----
    Private Type GDIPlusStartupInput
        GdiPlusVersion As Long
        DebugEventCallback As Long
        SuppressBackgroundThread As Long
        SuppressExternalCodecs As Long
    End Type
    
    ' ----==== GDIPlus Enums ====----
    Public Enum Status 'GDI+ Status
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
    
    ' ----==== GDI+ API Declarationen ====----
    Private Declare Function GdiplusStartup Lib "gdiplus" _
        (ByRef token As Long, ByRef lpInput As GDIPlusStartupInput, _
        Optional ByRef lpOutput As Any) As Status
    
    Private Declare Function GdiplusShutdown Lib "gdiplus" _
        (ByVal token As Long) As Status
    
    Private Declare Function GdipLoadImageFromFile Lib "gdiplus" _
        (ByVal FileName As Long, ByRef image As Long) As Status
    
    Private Declare Function GdipGetImageThumbnail Lib "gdiplus" _
        (ByVal image As Long, ByVal thumbWidth As Long, _
        ByVal thumbHeight As Long, ByRef thumbImage As Long, _
        ByVal callback As Long, ByVal callbackData As Long) _
        As Status
    
    Private Declare Function GdipGetImageDimension Lib "gdiplus" _
        (ByVal image As Long, ByRef Width As Single, _
        ByRef Height As Single) As Status
    
    Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" _
        (ByVal Bitmap As Long, ByRef hbmReturn As Long, _
        ByVal background As Long) As Status
    
    Private Declare Function GdipDisposeImage Lib "gdiplus" _
        (ByVal image As Long) As Status
    
    ' ----==== OLE API Declarations ====----
    Private Declare Sub OleCreatePictureIndirect Lib "oleaut32.dll" _
        (lpPictDesc As PICTDESC, riid As IID, ByVal fOwn As Boolean, _
        lplpvObj As Object)
    
    ' ----==== Variablen ====----
    Private GdipToken As Long
    Public GdipInitialized As Boolean
    Private Declare Function timeGetTime Lib "winmm.dll" () As Long


'------------------------------------------------------
' Funktion     : StartUpGDIPlus
' Beschreibung : Initialisiert GDI+ Instanz
' Übergabewert : GDI+ Version
' Rückgabewert : GDI+ Status
'------------------------------------------------------
Public Function StartUpGDIPlus(ByVal GdipVersion As Long) As Status
    ' Initialisieren der GDI+ Instanz
    Dim GdipStartupInput As GDIPlusStartupInput
    GdipStartupInput.GdiPlusVersion = GdipVersion
    StartUpGDIPlus = GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
End Function

'------------------------------------------------------
' Funktion     : ShutdownGDIPlus
' Beschreibung : Beendet die GDI+ Instanz
' Rückgabewert : GDI+ Status
'------------------------------------------------------
Public Function ShutdownGDIPlus() As Status
    ' Beendet GDI+ Instanz
    ShutdownGDIPlus = GdiplusShutdown(GdipToken)
End Function

'------------------------------------------------------
' Funktion     : Execute
' Beschreibung : Gibt im Fehlerfall die entsprechende GDI+ Fehlermeldung aus
' Übergabewert : GDI+ Status
' Rückgabewert : GDI+ Status
'------------------------------------------------------
Public Function Execute(ByVal lReturn As Status) As Status
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
' Übergabewert : GDI+ Status
' Rückgabewert : Fehlercode als String
'------------------------------------------------------
Private Function GdiErrorString(ByVal lError As Status) As String
    Dim s As String
    
    Select Case lError
    Case GenericError:              s = "Generic Error."
    Case InvalidParameter:          s = "Invalid Parameter."
    Case OutOfMemory:               s = "Out Of Memory."
    Case ObjectBusy:                s = "Object Busy."
    Case InsufficientBuffer:        s = "Insufficient Buffer."
    Case NotImplemented:            s = "Not Implemented."
    Case Win32Error:                s = "Win32 Error."
    Case WrongState:                s = "Wrong State."
    Case Aborted:                   s = "Aborted."
    Case FileNotFound:              s = "File Not Found."
    Case ValueOverflow:             s = "Value Overflow."
    Case AccessDenied:              s = "Access Denied."
    Case UnknownImageFormat:        s = "Unknown Image Format."
    Case FontFamilyNotFound:        s = "FontFamily Not Found."
    Case FontStyleNotFound:         s = "FontStyle Not Found."
    Case NotTrueTypeFont:           s = "Not TrueType Font."
    Case UnsupportedGdiplusVersion: s = "Unsupported Gdiplus Version."
    Case GdiplusNotInitialized:     s = "Gdiplus Not Initialized."
    Case PropertyNotFound:          s = "Property Not Found."
    Case PropertyNotSupported:      s = "Property Not Supported."
    Case Else:                      s = "Unknown GDI+ Error."
    End Select
    
    GdiErrorString = s
End Function

'------------------------------------------------------
' Funktion     : CreateThumbnailFromFile
' Beschreibung : Lädt ein Bilddatei per GDI+
' Übergabewert : FileName = Pfad\Dateiname der Bilddatei
'                Percent = Größe in Prozent (100% = 1:1)
' Rückgabewert : StdPicture Objekt
'------------------------------------------------------
Public Function CreateThumbnailFromFile(ByVal FileName As String, ByVal Percent As Long) As StdPicture
    
    Dim retStatus As Status
    Dim lBitmap As Long
    Dim lThumb As Long
    Dim hBitmap As Long
    Dim ImageWidth As Single
    Dim ImageHeight As Single
    Dim IW As Long
    Dim IH As Long
    
    ' Öffnet die Bilddatei in lBitmap
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
Private Function HandleToPicture(ByVal hGDIHandle As Long, _
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
    
    ' Rückgabe des Pictureobjekts
    Set HandleToPicture = oPicture
    
End Function


