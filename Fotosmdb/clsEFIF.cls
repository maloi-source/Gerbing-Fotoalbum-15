VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEXIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------------------------
'- EXIF Meta Tag reader
'- author: Chavdar Jordanov Änderungen zu GPS:Gerbing 05.06.2013
'- based on the Exif format description Exif Version 2.3 mit GPS.pdf (CIPA DC-008-Translation.2010)
'------------------------------------------------------------------------------------------------------------
Option Explicit
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    
    Private Type tIDF 'main structure, containg tag parameters
        ID As Long
        Name As String
        Format As Long
        Length As Long
        Data As Variant
        StringData As String
    End Type
    
    Public Enum enTAG 'enumerator containing the most common metatag names and their respective IDs
        GPSVersionID = 0
        GPSLatitudeRef = 1
        GPSLatitude = 2
        GPSLongitudeRef = 3
        GPSLongitude = 4
        GPSAltitudeRef = 5
        GPSAltitude = 6
        GPSTimeStamp = 7
        GPSSatellites = 8
        GPSStatus = 9
        GPSMeasureMode = 10
        GPSDOP = 11
        GPSSpeedRef = 12
        GPSSpeed = 13
        GPSTrackRef = 14
        GPSTrack = 15
        GPSImgDirectionRef = 16
        GPSImgDirection = 17
        GPSMapDatum = 18
        GPSDestLatitudeRef = 19
        GPSDestLatitude = 20
        GPSDestLongitudeRef = 21
        GPSDestLongitude = 22
        GPSDestBearingRef = 23
        GPSDestBearing = 24
        GPSDestDistanceRef = 25
        GPSDestDistance = 26
        GPSProcessingMethod = 27
        GPSAreaInformation = 28
        GPSDateStamp = 29
        GPSDifferential = 30
        GPSHPositioningError = 31
        ImageWidth = 256                                                'Gerbing 13.06.2013
        ImageHeight = 257                                               'Gerbing 13.06.2013
        ImageDescription = 270
        Make = 271
        Model = 272
        Orientation = 274
        XResolution = 282
        YResolution = 283
        ResolutionUnit = 296
        Software = 305
        DateTime = 306
        Artist = 315
        WhitePoint = 318
        PrimaryChromaticities = 319
        YCbCrCoefficients = 529
        YCbCrPositioning = 531
        ReferenceBlackWhite = 532
        Rating = 18246                                                  'Gerbing 19.06.2013
        RatingPercent = 18249                                           'Gerbing 19.06.2013
        Copyright = 33432
        ExifOffset = 34665
        Exposuretime = 33434
        FNumber = 33437
        ICCProfile = 34675                                              'Gerbing 13.06.2013
        ExposureProgram = 34850
        ISOSpeedRatings = 34855
        SensitivityType = 34864                                         'Gerbing 13.06.2013
        ExifVersion = 36864
        DateTimeOriginal = 36867
        DateTimeDigitized = 36868
        ComponentsConfiguration = 37121
        CompressedBitsPerPixel = 37122
        ShutterSpeedValue = 37377
        ApertureValue = 37378
        BrightnessValue = 37379
        ExposureBiasValue = 37380
        MaxApertureValue = 37381
        SubjectDistance = 37382
        MeteringMode = 37383
        LightSource = 37384
        Flash = 37385
        FocalLength = 37386
        MakerNote = 37500
        UserComment = 37510
        SubsecTime = 37520
        SubsecTimeOriginal = 37521
        SubsecTimeDigitized = 37522
        XPTitle = 40091                                                 'Gerbing 16.06.2013
        XPComment = 40092
        XPAuthor = 40093
        XPKeywords = 40094
        XPSubject = 40095                                               'Gerbing 16.06.2013
        FlashPixVersion = 40960
        ColorSpace = 40961
        ExifImageWidth = 40962
        ExifImageHeight = 40963
        RelatedSoundFile = 40964
        ExifInteroperabilityOffset = 40965
        FocalPlaneXResolution = 41486
        FocalPlaneYResolution = 41487
        FocalPlaneResolutionUnit = 41488
        ExposureIndex = 41493
        SensingMethod = 41495
        FileSource = 41728
        SceneType = 41729
        CFAPattern = 41730
        CustomRendered = 41985
        ExposureMode = 41986
        WhiteBalance = 41987
        DigitalZoomRatio = 41988
        SceneCaptureType = 41990
        GainControl = 41991
        Contrast = 41992
        Saturation = 41993
        Sharpness = 41994
        SubjectDistanceRange = 41996
        ImageUniqueID = 42016
        Padding = 59932
    End Enum
    
    Private Enum enFormat
        unsignedByte = 1
        asciiString = 2
        unsignedShort = 3
        unsignedLong = 4
        unsignedRationale = 5
        signedByte = 6
        undefined = 7
        signedShort = 8
        signedLong = 9
        signedRationale = 10
        singleFloat = 11
        doubleFloat = 12
    End Enum
    
    Private m_ImageFile As String 'path to the file on the disk
    Private m_Intel As Boolean    'flag containing the byte alignment of the record
    Private m_IniFile As String   'path to the exif.ini file containing description of the codes
    
    Private DataLen 'array containing the byte length of each data format
    Private IDName  'array containing tag names
    Private IDNo    'array containing tag IDs
    Private bParsed As Byte 'flag that the file has been parsed succesfuly
    Private IFD() As tIDF 'array containing all metatags
    Private GPSOffset As Long                                       'Gerbing 05.06.2013
    Private NextIFD As Long                                         'Gerbing 05.06.2013
    '======== Interface part ==========

'-- file containing the jpeg image --
Public Property Get ImageFile() As Variant
    ImageFile = m_ImageFile
End Property

Public Property Let ImageFile(ByVal vNewValue As Variant)
    m_ImageFile = vNewValue
    bParsed = ReadMetaInfo(m_ImageFile)
End Property

'-- returns the byte alignment order for the file --
Public Property Get IntelByteAlignment() As Boolean
    IntelByteAlignment = m_Intel
End Property

'-- method, which returns the numeric and string values for a single metatag ---
'kommt überhauptnicht dran              'gerbing
Public Function MetaInfo(ByVal l_ID As enTAG, ByRef StringData As String) As Long
    Dim i As Integer
    
    If bParsed = 0 Then
        For i = 1 To UBound(IFD)
            If IFD(i).ID = l_ID Then
                StringData = IFD(i).StringData
                MetaInfo = IFD(i).Data
                Exit Function
            End If
        Next i
        StringData = "Tag " + CStr(l_ID) + " not found."
    ElseIf bParsed = 1 Then
        'Err.Raise 10, "MetaInfo", "File is not in EXIF format."    'Gerbing auskommentiert
    ElseIf bParsed = 2 Then
        'Err.Raise 11, "MetaInfo", "Error parsing the file."        'Gerbing auskommentiert
    End If
End Function

'-- returns Exif tag name based on its ID --
Function GetTagName(ByVal lID As enTAG) As String
    Dim i As Integer
    
    For i = 0 To UBound(IDNo)
        If lID = IDNo(i) Then
            GetTagName = IDName(i)
            Exit Function
        End If
    Next
    'tag name unknown; return tag ID
    GetTagName = "unknown tag " + CStr(lID)
End Function

'--- Lists all metatags found in the header ---
Function ListInfo() As String
    Dim i As Integer
    
    On Error Resume Next                                        'Gerbing 19.02.2011
    i = UBound(IFD)
    If ERR.Number <> 0 Then
        ListInfo = "File is not in EXIF format."
        Exit Function
    End If
    On Error GoTo 0
    '-------------------------------------------
    If bParsed = 0 Then
        If UBound(IFD) > 200 Then                               'Gerbing 03.06.2013
            For i = 1 To 200
            If IFD(i).Name <> "" Then
                If IFD(i).Name <> "ExifInteroperabilityOffset" Then
                    ListInfo = ListInfo + IFD(i).Name + ": " + IFD(i).StringData + vbCrLf
                End If
            End If
            Next i
        Else
            For i = 1 To UBound(IFD)
            If IFD(i).Name <> "" Then
                If IFD(i).Name <> "ExifInteroperabilityOffset" Then
                    ListInfo = ListInfo + IFD(i).Name + ": " + IFD(i).StringData + vbCrLf
                End If
            End If
            Next i
        End If
    ElseIf bParsed = 1 Then
        ListInfo = "File is not in EXIF format."
    ElseIf bParsed = 2 Then
        ListInfo = "Could not open the file."
    End If
End Function

'========= PARSING FUNCTIONS ==========

'-- parses the jpeg header and extracts all Exif information from it --
Private Function ReadMetaInfo(sFileName As String) As Integer
    Dim sJPEG_Header As String, b() As Byte
    Dim lPos As Long, Offset As Long, HeaderStart As Long
    Dim GPSLatitudeStart As Long                                                                        'Gerbing 08.04.2019
    Dim GPSLongitudeStart As Long                                                                       'Gerbing 08.04.2019
    Dim strLatitude As String                                                                           'Gerbing 08.04.2019
    Dim strLongitude As String                                                                          'Gerbing 08.04.2019
    Dim i As Integer
    Dim NoOfRecs As Integer
    
    On Error GoTo ErrRead
    GPSOffset = 0
    sJPEG_Header = MyReadFile(sFileName) 'may be changed to reflect the actual header size              'Gerbing 04.03.2013
    If sJPEG_Header = "" Then
        ERR.Raise 2, "ReadMetaInfo", "File not found."
    Else
        'Suchen im XMP-Abschnitt                                                                        'Gerbing 08.04.2019
        gstrLatXMP = ""                                                                                 'Gerbing 08.04.2019
        gstrLongXMP = ""                                                                                'Gerbing 08.04.2019
        GPSLatitudeStart = InStrB(1, sJPEG_Header, "exif:GPSLatitude", vbTextCompare)                   'Gerbing 08.04.2019
        GPSLongitudeStart = InStrB(1, sJPEG_Header, "exif:GPSLongitude", vbTextCompare)                 'Gerbing 08.04.2019
        If GPSLatitudeStart <> 0 Then
            strLatitude = Mid(sJPEG_Header, GPSLatitudeStart \ 2, 50) 'zB<exif:GPSLatitude>50,38.7309456N</exif:            'Gerbing 08.04.2019
            lPos = InStr(1, strLatitude, "</exif", vbTextCompare)
            gstrLatXMP = Mid(strLatitude, 19, lPos - 19)
            If StrComp(Mid(gstrLatXMP, 1, 3), "0,0") = 0 Then
                gstrLatXMP = ""
            End If
        End If
        If GPSLongitudeStart <> 0 Then
            strLongitude = Mid(sJPEG_Header, GPSLongitudeStart \ 2, 50) 'zB<exif:GPSLongitude>11,53.9826786E</exif:         'Gerbing 08.04.2019
            lPos = InStr(1, strLongitude, "</exif", vbTextCompare)
            gstrLongXMP = Mid(strLongitude, 20, lPos - 20)
            If StrComp(Mid(gstrLongXMP, 1, 3), "0,0") = 0 Then
                gstrLongXMP = ""
            End If
        End If
        HeaderStart = InStr(1, sJPEG_Header, "Exif" + Chr(0), vbBinaryCompare) 'start of EXIF header
        If HeaderStart = 0 Then ReadMetaInfo = 1: Exit Function
        HeaderStart = HeaderStart + 6 'start of data
        lPos = HeaderStart
        m_Intel = Mid(sJPEG_Header, lPos, 2) = "II" 'byte alignment
        lPos = lPos + 4
        Offset = BVal(sJPEG_Header, lPos, 4, m_Intel) 'offset to the first IFD
        ReDim IFD(0)
        'parse the main IFD directory and get the offset to the IFDSubDir
        Offset = ParseIDF(sJPEG_Header, HeaderStart, Offset, ExifOffset)                                '34665
        Offset = ParseIDF(sJPEG_Header, HeaderStart, Offset, ExifInteroperabilityOffset)                '40965
        
'        If NextIFD <> 0 Then
'            Offset = ParseIDF(sJPEG_Header, HeaderStart, NextIFD, ExifInteroperabilityOffset)
'        End If
        
        If GPSOffset <> 0 Then
            Offset = ParseIDF(sJPEG_Header, HeaderStart, GPSOffset, ExifInteroperabilityOffset)         'Gerbing 05.06.2013
        End If
    End If
ExitRead:
    Exit Function
ErrRead:
'    Dim S                                      'gerbing auskommentiert
'    S = Err.Description
'    Err.Raise 1, "ReadMetaInfo", S
    'ReadMetaInfo = 2                           'Gerbing 07.05.2007
End Function

'--- Parses a single EXIF directory record (IFD), returns an offset to the next IFD ---
Private Function ParseIDF(sJPEG_Header As String, HeaderStart As Long, Offset As Long, ClosingTag As enTAG) As Long
    Dim lPos As Long, NoOfRecs As Long, i As Integer, u
    
    lPos = HeaderStart + Offset 'get the starting offset position
    NoOfRecs = BVal(sJPEG_Header, lPos, 2, m_Intel) 'get the No of records in the IFD
    If NoOfRecs > 200 Then
        NoOfRecs = 200                                                      'Gerbing 03.06.2013
    End If
    u = UBound(IFD)
    If u > 100 Then
        u = 100                                                              'Gerbing 03.06.2013
    End If
    ReDim Preserve IFD(u + NoOfRecs) 'redimension the IFD array
    lPos = lPos + 2
    On Error GoTo 0                             'Gerbing 07.05.2007
    'begin retrieving the tags
    For i = u + 1 To u + NoOfRecs
        If BVal(sJPEG_Header, lPos, 2, m_Intel) = 34853 Then                'ID for GPS IFD Offset
            GPSOffset = BVal(sJPEG_Header, lPos + 8, 4, m_Intel)            'Gerbing 05.06.2013
            GoTo GotoNext12
        End If
        If BVal(sJPEG_Header, lPos, 2, m_Intel) = 34665 Then                'ID for Exif IFD Offset
            ParseIDF = BVal(sJPEG_Header, lPos + 8, 4, m_Intel)  'return offset to the next sub IFD
            GoTo GotoNext12
        End If
        IFD(i).ID = BVal(sJPEG_Header, lPos, 2, m_Intel) 'first 2 bytes contain the tag ID
        IFD(i).Format = BVal(sJPEG_Header, lPos + 2, 2, m_Intel) 'next 2 bytes contain the tag data format
        IFD(i).Length = BVal(sJPEG_Header, lPos + 4, 4, m_Intel) 'next 4 bytes contain the No of data components
        IFD(i).Data = BVal(sJPEG_Header, lPos + 8, 4, m_Intel) 'next 4 bytes contain the data or an offset to the data
        IFD(i).Name = GetTagName(IFD(i).ID) 'retrieve the tag name in human readable format
        If IFD(i).Name = "ICCProfile" Then                                      'Gerbing 13.06.2013
            IFD(i).StringData = "not supported"
            GoTo GotoNext12
        End If
        If IFD(i).Name = "Padding" Then                                         'Gerbing 16.06.2013
            IFD(i).StringData = "not supported"
            GoTo GotoNext12
        End If
        If IFD(i).Length = 0 Then                                               'Gerbing 02.07.2014
            GoTo GotoNext12
        End If

'        If IFD(i).Name = "ImageUniqueID" Then                                  'Gerbing 13.06.2013 zum Test zb was kommt bei ImageUniqueID
'            '
'            IFD(i).Name = "ImageUniqueID"
'        End If
        Call GetStringData(sJPEG_Header, IFD(i), HeaderStart, lPos) 'retrieve the data as a string
GotoNext12:
        lPos = lPos + 12 'get the offset to the next tag
    Next
    If ClosingTag = ExifOffset Then
        NextIFD = BVal(sJPEG_Header, lPos, 4, m_Intel) 'next 4 bytes contain the offset to the next IFD
    End If
End Function

'-- returns formatted string data from an IFD entry ---
Private Sub GetStringData(sJPEG_Header As String, IFD As tIDF, OffsetStart As Long, CurrPosition As Long)
    Dim vData As String, X As Variant, sFmt As String
    Dim PixOrder, i As Integer
    Dim pos As Long                                                                 'Gerbing 03.10.2016
    Dim pos1 As Long
    
    'extract the data
    If IFD.Length * DataLen(IFD.Format) > 4 Then 'IFD data contains offset to real data
        vData = Mid(sJPEG_Header, OffsetStart + IFD.Data, IFD.Length * DataLen(IFD.Format))
    Else 'IFD record contains the data itself
        vData = Mid(sJPEG_Header, CurrPosition + 8, 4)
    End If
    'check the data format and do some conversions if necessary
    Select Case IFD.Format
    Case enFormat.asciiString                                                       'ascii=2
        IFD.StringData = vData
        Select Case IFD.ID
        Case enTAG.ImageUniqueID
            IFD.StringData = vData
        End Select
        'hier ConvertFromUTF8                                                       'Gerbing 14.11.2015
        IFD.StringData = FromUTF8String(Mid(vData, 1))                              'Gerbing 14.11.2015
    Case enFormat.unsignedRationale                                                 'unsigned rational=5
        Select Case IFD.ID
        Case enTAG.ApertureValue, enTAG.MaxApertureValue
            'convert to aperture ratio value
            X = UnsignedRational(vData, IFD.Data)
            IFD.StringData = Format(Sqr(2) ^ IFD.Data, "0.0")
        Case enTAG.FNumber, enTAG.GPSDOP, enTAG.GPSSpeed, enTAG.GPSTrack, enTAG.GPSImgDirection, enTAG.GPSDestBearing, _
                enTAG.GPSDestDistance, enTAG.GPSHPositioningError
            X = UnsignedRational(vData, IFD.Data)
            IFD.StringData = Format(IFD.Data, "0.0")
        Case enTAG.FocalPlaneXResolution, enTAG.FocalPlaneYResolution, enTAG.XResolution, enTAG.YResolution, enTAG.FocalLength, _
                enTAG.SubjectDistance, enTAG.DigitalZoomRatio
            X = UnsignedRational(vData, IFD.Data)
            IFD.StringData = Format(IFD.Data, "0.0")
        Case enTAG.Exposuretime
            X = UnsignedRationalExposure(vData, IFD.Data)
            IFD.StringData = "1/" & X
        Case enTAG.GPSLatitude, enTAG.GPSLongitude, enTAG.GPSDestLatitude, enTAG.GPSDestLongitude
            X = UnsignedRational(Mid(vData, 1, 8), IFD.Data)                            'Grad
            IFD.StringData = Format(IFD.Data, "0") & "."                                'Gerbing 02.09.2016
            X = UnsignedRational(Mid(vData, 9, 8), IFD.Data)
            IFD.StringData = IFD.StringData + IFD.Data / 60                             'Minuten
            X = UnsignedRational(Mid(vData, 17, 8), IFD.Data)
            IFD.StringData = IFD.StringData + IFD.Data / 3600                           'Sekunden
            'begrenzen auf max 8 Stellen nach dem Komma                                 'Gerbing 03.10.2016
            pos = InStr(1, IFD.StringData, ",")                                         'suche Komma
            pos1 = Len(IFD.StringData)                                                  'Gesamtlänge
            If pos1 > pos + 8 Then
                IFD.StringData = Mid(IFD.StringData, 1, pos + 8)
            End If
        Case enTAG.GPSAltitude
            X = UnsignedRational(vData, IFD.Data)
            IFD.StringData = Format(IFD.Data, "0000.0")
        Case enTAG.GPSTimeStamp
            X = UnsignedRational(Mid(vData, 1, 8), IFD.Data)
            IFD.StringData = Format(IFD.Data, "0") & ":"
            X = UnsignedRational(Mid(vData, 9, 8), IFD.Data)
            IFD.StringData = IFD.StringData & Format(IFD.Data, "0") & ":"
            X = UnsignedRational(Mid(vData, 17, 8), IFD.Data)
            IFD.StringData = IFD.StringData & Format(IFD.Data, "0")
        Case Else
            If IFD.Length < 2 Then
                IFD.StringData = UnsignedRational(vData, IFD.Data)
            Else
                For i = 1 To IFD.Length
                    Call UnsignedRational(Mid(vData, (i - 1) * 8 + 1, 4), X)
                    IFD.StringData = "/" + IFD.StringData + Format(X, "0.000")
                Next
                IFD.StringData = Mid(IFD.StringData, 2)
            End If
        End Select
    Case enFormat.signedRationale                                                   'signed rational=10
        Select Case IFD.ID
        Case enTAG.ShutterSpeedValue
            'convert to shutter speed value
            X = SignedRational(vData, IFD.Data)
            IFD.StringData = Format(2 ^ IFD.Data, "0")
            IFD.StringData = "1/" & IFD.StringData
        Case enTAG.ExposureBiasValue
            X = SignedRational(vData, IFD.Data)
            IFD.StringData = Format(IFD.Data, "0.0")
        Case Else
            If IFD.Length < 2 Then
                IFD.StringData = SignedRational(vData, IFD.Data)
            Else
                For i = 1 To IFD.Length
                    Call SignedRational(Mid(vData, (i - 1) * 8 + 1, 4), X)
                    IFD.StringData = "/" + IFD.StringData + Format(X, "0.000")
                Next
                IFD.StringData = Mid(IFD.StringData, 2)
            End If
        End Select
    Case enFormat.undefined                                                         'undefined=7
        Select Case IFD.ID
        Case enTAG.MakerNote
            IFD.StringData = ExtractTextOnly(vData)
        Case enTAG.ComponentsConfiguration
            PixOrder = Array("", "Y", "Cb", "Cr", "R", "G", "B")
            For i = 1 To 4
                IFD.StringData = IFD.StringData + PixOrder(Asc(Mid(vData, i, 1)))
            Next i
        Case enTAG.ExifVersion, enTAG.FlashPixVersion
            IFD.StringData = CStr(Val(left(vData, 2))) + "." + Right(vData, 2)
        Case enTAG.FileSource
            IFD.StringData = CStr(BVal(vData, 1, DataLen(IFD.Format), m_Intel))
            Select Case IFD.StringData
            Case "0"
                IFD.StringData = "unknown"
            Case "1"
                IFD.StringData = "scanner of transparent type"
            Case "2"
                IFD.StringData = "scanner of reflex type"
            Case "3"
                IFD.StringData = "digital camera"
            End Select
        Case enTAG.SceneType
            IFD.StringData = CStr(BVal(vData, 1, DataLen(IFD.Format), m_Intel))
            Select Case IFD.StringData
            Case "1"
                IFD.StringData = "directly photographed image"
            End Select
        Case enTAG.UserComment
            IFD.StringData = ExtractTextOnly(vData)                                         'Gerbing 13.06.2013
        Case Else
            IFD.StringData = vData
        End Select
    Case enFormat.unsignedShort, enFormat.signedByte, enFormat.signedLong, enFormat.signedShort
                                                                            'unsignedshort=3 signedbyte=6, signedshort=8 & signedlong=9
        Select Case IFD.ID
        Case enTAG.GainControl
            IFD.StringData = CStr(BValS(vData, 1, DataLen(IFD.Format), m_Intel))
            Select Case IFD.StringData
            Case "0"
                IFD.StringData = "none"
            Case "1"
                IFD.StringData = "low gain up"
            Case "2"
                IFD.StringData = "high gain up"
            Case "3"
                IFD.StringData = "low gain down"
            Case "4"
                IFD.StringData = "high gain down"
            End Select
        Case enTAG.Contrast
            IFD.StringData = CStr(BValS(vData, 1, DataLen(IFD.Format), m_Intel))
            Select Case IFD.StringData
            Case "0"
                IFD.StringData = "normal"
            Case "1"
                IFD.StringData = "soft"
            Case "2"
                IFD.StringData = "hard"
            End Select
        Case enTAG.Saturation
            IFD.StringData = CStr(BValS(vData, 1, DataLen(IFD.Format), m_Intel))
            Select Case IFD.StringData
            Case "0"
                IFD.StringData = "normal"
            Case "1"
                IFD.StringData = "low"
            Case "2"
                IFD.StringData = "high"
            End Select
        Case enTAG.Orientation
            IFD.StringData = CStr(BValS(vData, 1, DataLen(IFD.Format), m_Intel))
            Select Case IFD.StringData
            Case "1"
                IFD.StringData = "horizontal normal"
            Case Else
                IFD.StringData = "mirrored or turned"
            End Select
        Case enTAG.YCbCrPositioning
            IFD.StringData = CStr(BValS(vData, 1, DataLen(IFD.Format), m_Intel))
            Select Case IFD.StringData
            Case "1"
                IFD.StringData = "centered"
            Case "2"
                IFD.StringData = "co-sited"
            End Select
        Case enTAG.ResolutionUnit
            IFD.StringData = CStr(BValS(vData, 1, DataLen(IFD.Format), m_Intel))
            Select Case IFD.StringData
            Case "2"
                IFD.StringData = "inches"
            Case "3"
                IFD.StringData = "centimeters"
            End Select
        Case enTAG.MeteringMode
            IFD.StringData = CStr(BValS(vData, 1, DataLen(IFD.Format), m_Intel))
            Select Case IFD.StringData
            Case "0"
                IFD.StringData = "unknown"
            Case "1"
                IFD.StringData = "average"
            Case "2"
                IFD.StringData = "CenterWeightedAverage"
            Case "3"
                IFD.StringData = "spot"
            Case "4"
                IFD.StringData = "MultiSpot"
            Case "5"
               IFD.StringData = "pattern"
            Case "6"
                IFD.StringData = "partial"
            Case "255"
                IFD.StringData = "other"
            End Select
        Case enTAG.ColorSpace
            IFD.StringData = CStr(BValS(vData, 1, DataLen(IFD.Format), m_Intel))
            Select Case IFD.StringData
            Case "1"
                IFD.StringData = "sRGB"
            Case Else
                IFD.StringData = "uncalibrated"
            End Select
        Case enTAG.SensingMethod
            IFD.StringData = CStr(BValS(vData, 1, DataLen(IFD.Format), m_Intel))
            Select Case IFD.StringData
            Case "1"
                IFD.StringData = "not defined"
            Case "2"
                IFD.StringData = "one-chip color area sensor"
            Case "3"
                IFD.StringData = "two-chip color area sensor"
            Case "4"
                IFD.StringData = "three-chip color area sensor"
            Case "5"
                IFD.StringData = "color sequential area sensor"
            Case "7"
               IFD.StringData = "trilinear sensor"
            Case "8"
                IFD.StringData = "color sequential linear sensor"
            End Select
        Case enTAG.SceneCaptureType
            IFD.StringData = CStr(BValS(vData, 1, DataLen(IFD.Format), m_Intel))
            Select Case IFD.StringData
            Case "0"
                IFD.StringData = "standard"
            Case "1"
                IFD.StringData = "landscape"
            Case "2"
                IFD.StringData = "portrait"
            Case "3"
                IFD.StringData = "night scene"
            End Select
        Case enTAG.ExposureMode
            IFD.StringData = CStr(BValS(vData, 1, DataLen(IFD.Format), m_Intel))
            Select Case IFD.StringData
            Case "0"
                IFD.StringData = "auto"
            Case "1"
                IFD.StringData = "manual"
            Case "2"
                IFD.StringData = "auto bracket"
            End Select
        Case enTAG.WhiteBalance
            IFD.StringData = CStr(BValS(vData, 1, DataLen(IFD.Format), m_Intel))
            Select Case IFD.StringData
            Case "0"
                IFD.StringData = "auto"
            Case "1"
                IFD.StringData = "manual"
            End Select
        Case enTAG.ISOSpeedRatings
            IFD.StringData = CStr(BValS(vData, 1, DataLen(IFD.Format), m_Intel))
            IFD.Name = "ISO"
        Case enTAG.SensitivityType
            IFD.StringData = CStr(BValS(vData, 1, DataLen(IFD.Format), m_Intel))
            Select Case IFD.StringData
            Case "0"
                IFD.StringData = "unknown"
            Case "1"
                IFD.StringData = "standard output sensitivity"
            Case "2"
                IFD.StringData = "recommended exposure index"
            Case "3"
                IFD.StringData = "ISO speed"
            Case "4"
                IFD.StringData = "standard output sensitivity and recommended exposure index"
            Case "5"
                IFD.StringData = "standard output sensitivity and ISO speed"
            Case "6"
                IFD.StringData = "recommended exposure index and ISO speed"
            Case "7"
                IFD.StringData = "standard output sensitivity and recommended exposure index and ISO speed"
            End Select
        Case Else
            IFD.StringData = CStr(BValS(vData, 1, DataLen(IFD.Format), m_Intel))
        End Select
    Case enFormat.unsignedByte                                                      'unsigned byte=1
        Select Case IFD.ID
        Case enTAG.GPSVersionID
            'IFD.StringData = CStr(BVal(vData, 1, DataLen(IFD.Format), m_Intel))
            IFD.StringData = CStr(BVal(vData, 1, 1, m_Intel)) & "."
            IFD.StringData = IFD.StringData & CStr(BVal(vData, 2, 1, m_Intel)) & "."
            IFD.StringData = IFD.StringData & CStr(BVal(vData, 3, 1, m_Intel)) & "."
            IFD.StringData = IFD.StringData & CStr(BVal(vData, 3, 1, m_Intel))
        Case enTAG.XPTitle, enTAG.XPSubject, enTAG.XPKeywords, enTAG.XPComment, enTAG.XPAuthor              'Gerbing 16.06.2013
            IFD.StringData = StrConv(vData, vbFromUnicode)
        End Select
    Case Else 'all other data types
        IFD.StringData = CStr(BVal(vData, 1, DataLen(IFD.Format), m_Intel))
    End Select
    IFD.StringData = Replace(IFD.StringData, Chr(0), "") 'remove null characters
    'try to read description of data from the ini file
    vData = GetProfileString(m_IniFile, IFD.Name, IFD.StringData)
    sFmt = GetProfileString(m_IniFile, IFD.Name, "Format")
    'if description is found, use it instead of raw data
    If vData <> "" Then IFD.StringData = vData
    If sFmt <> "" Then IFD.StringData = Replace(sFmt, "@X", IFD.StringData)
End Sub

'======== byte hadling functions ==========

'--- returns calculated unsigned value of a byte sequence ---
Private Function BVal(sData As String, start As Long, ByVal Length As Integer, m_Intel As Boolean) As Variant
    Dim i As Long, st As Long, en As Long, step As Integer, n As Long
    
    If Not m_Intel Then 'Motorola byte alignment
        st = start + Length - 1
        en = start
        step = -1
    Else                 'Intel byte alignment
        st = start
        en = start + Length - 1
        step = 1
    End If
    For i = st To en Step step
        BVal = BVal + Asc(Mid(sData, i, 1)) * 256 ^ n
        n = n + 1
    Next
End Function

'--- returns calculated signed value of a byte sequence ---
Private Function BValS(sData As String, start As Long, ByVal Length As Integer, m_Intel As Boolean) As Variant
    Dim dn As Variant
    
    dn = 2 ^ (8 * Length - 1)
    BValS = BVal(sData, start, Length, m_Intel)
    If BValS > (dn - 1) Then BValS = Not (BValS - dn)
End Function

'--- returns a string containing an unsigned rational value in the format Numerator/Denumerator for Exposuretime
'- also calculates its numeric value
Private Function UnsignedRationalExposure(ByVal vData As String, ByRef NumValue As Variant) As String
    Dim num As Double, Denum As Double, Rational As Double
    
    num = BVal(vData, 1, 4, m_Intel)
    Denum = BVal(vData, 5, 4, m_Intel)
    If num <> 0 Then
        Rational = Denum / num
    End If
    UnsignedRationalExposure = CStr(Rational)
End Function

'--- returns a string containing an unsigned rational value in the format Numerator/Denumerator;
'- also calculates its numeric value
Private Function UnsignedRational(ByVal vData As String, ByRef NumValue As Variant) As String
    Dim num As Double, Denum As Double, Rational As Double
    num = BVal(vData, 1, 4, m_Intel)
    Denum = BVal(vData, 5, 4, m_Intel)
    If Denum <> 0 Then NumValue = num / Denum
    UnsignedRational = CStr(num) + "/" + CStr(Denum)
End Function

'--- returns a string containing a signed rational value in the format Numerator/Denumerator;
'- also calculates its numeric value
Private Function SignedRational(ByVal vData As String, ByRef NumValue As Variant) As String
    Dim num As Double, Denum As Double, Rational As Double
    
    num = BValS(vData, 1, 4, m_Intel)
    Denum = BValS(vData, 5, 4, m_Intel)
    If Denum <> 0 Then NumValue = num / Denum
    SignedRational = CStr(num) + "/" + CStr(Denum)
End Function

Private Function MyReadFile(ByVal sFilePath As String) As String                        'Gerbing 04.03.2013
    Dim hHandle As Long
    Dim imageData() As Byte
    Dim imageString As String
    Dim bytesRead As Long

    hHandle = GetFileHandle(sFilePath, True)                             'true=read     'versteht unicode filename
    If hHandle <> INVALID_HANDLE_VALUE Then
        If hHandle Then
            bytesRead = GetFileSize(hHandle, ByVal 0&)
            If bytesRead Then
                ReDim imageData(0 To bytesRead - 1)
                ReadFile hHandle, imageData(0), bytesRead, bytesRead, ByVal 0&
                 If bytesRead > UBound(imageData) Then
                    imageString = ""
                    imageString = StrConv(imageData, vbUnicode)
                End If
            End If
            CloseHandle hHandle
        End If
    End If
    MyReadFile = imageString
End Function

'--- Reads a value from an .INI file ---
Private Function GetProfileString(ByVal sFile As String, ByVal sSection As String, ByVal sKey As String, Optional ByVal DefaultValue = "") As String
    Dim sTmp As String, X As Long
    
    Const StringSize = 1024
    sTmp = Space$(StringSize)
    X = GetPrivateProfileString(sSection, sKey, "", sTmp, StringSize, sFile)
    sTmp = Trim$(sTmp)
    sTmp = left(sTmp, Len(sTmp) - 1)
    If sTmp = "" Then sTmp = DefaultValue
    GetProfileString = sTmp
End Function

'--- Returns only the ascii characters from a byte sequence ---
Private Function ExtractTextOnly(S As String) As String
    Dim i As Integer, l As Integer, c As String * 1
    
    l = Len(S)
    For i = 1 To l
        c = Mid(S, i, 1)
        If Asc(c) > 31 And Asc(c) < 128 Then ExtractTextOnly = ExtractTextOnly + c
    Next i
End Function

'--- initializes arrays and other variables ---
Private Sub Class_Initialize()
    m_IniFile = AppPath
    If Right(m_IniFile, 1) <> "\" Then m_IniFile = m_IniFile + "\"
    m_IniFile = m_IniFile + "exif.ini"
    DataLen = Array(0, 1, 1, 2, 4, 8, 1, 1, 2, 4, 8, 4, 8)
    IDNo = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, _
        256, 257, 270, 271, 272, 274, 282, 283, 296, 305, 306, 315, 318, 319, 529, 531, 532, 18246, 18249, 33432, 34665, _
        33434, 33437, 34675, 34850, 34855, 34864, 36864, 36867, 36868, 37121, 37122, 37377, 37378, 37379, 37380, 37381, _
        37382, 37383, 37384, 37385, 37386, 37500, 37510, 37520, 37521, 37522, _
        40091, 40092, 40093, 40094, 40095, _
        40960, 40961, 40962, 40963, 40964, 40965, _
        41486, 41487, 41488, 41493, 41495, 41728, 41729, 41730, _
        41985, 41986, 41987, 41988, 41990, 41991, 41992, 41993, 41994, 41996, 42016, _
        59932)
    IDName = Array("GPSVersionID", "GPSLatitudeRef", "GPSLatitude", "GPSLongitudeRef", "GPSLongitude", "GPSAltitudeRef", "GPSAltitude", "GPSTimeStamp", "GPSSatellites", "GPSStatus", "GPSMeasureMode", "GPSDOP", "GPSSpeedRef", "GPSSpeed", "GPSTrackRef", "GPSTrack", "GPSImgDirectionRef", _
        "GPSImgDirection", "GPSMapDatum", "GPSDestLatitudeRef", "GPSDestLatitude", "GPSDestLongitudeRef", "GPSDestLongitude", "GPSDestBearingRef", "GPSDestBearing", "GPSDestDistanceRef", "GPSDestDistance", "GPSProcessingMethod", "GPSAreaInformation", "GPSDateStamp", "GPSDifferential", "GPSHPositioningError", _
        "ImageWidth", "ImageHeight", "ImageDescription", "Make", "Model", "Orientation", "XResolution", "YResolution", "ResolutionUnit", "Software", "DateTime", "Artist", "WhitePoint", "PrimaryChromaticities", "YCbCrCoefficients", "YCbCrPositioning", "ReferenceBlackWhite", "Rating", "RatingPercent", "Copyright", "ExifOffset", _
        "ExposureTime", "FNumber", "ICCProfile", "ExposureProgram", "ISOSpeedRatings", "SensitivityType", "ExifVersion", "DateTimeOriginal", "DateTimeDigitized", "ComponentsConfiguration", "CompressedBitsPerPixel", "ShutterSpeedValue", "ApertureValue", "BrightnessValue", "ExposureBiasValue", "MaxApertureValue", _
        "SubjectDistance", "MeteringMode", "LightSource", "Flash", "FocalLength", "MakerNote", "UserComment", "SubsecTime", "SubsecTimeOriginal", "SubsecTimeDigitized", _
        "XPTitle", "XPComment", "XPAuthor", "XPKeywords", "XPSubject", _
        "FlashPixVersion", "ColorSpace", "ExifImageWidth", "ExifImageHeight", "RelatedSoundFile", "ExifInteroperabilityOffset", _
        "FocalPlaneXResolution", "FocalPlaneYResolution", "FocalPlaneResolutionUnit", "ExposureIndex", "SensingMethod", "FileSource", "SceneType", "CFAPattern", _
        "CustomRendered", "ExposureMode", "WhiteBalance", "DigitalZoomRatio", "SceneCaptureType", "GainControl", "Contrast", "Saturation", "Sharpness", "SubjectDistanceRange", "ImageUniqueID", _
        "Padding")
End Sub

Function FromUTF8String(ByVal S As String) As String                                        'Gerbing 14.11.2015
   Dim i As Integer, b(2) As Byte
   
   i = 1
   S = S & Chr(0) & Chr(0)
   Do While i <= Len(S) - 2
      b(0) = Asc(Mid(S, i, 1))
      b(1) = Asc(Mid(S, i + 1, 1))
      b(2) = Asc(Mid(S, i + 2, 1))
      If (b(0) And &HE0) = &HE0 Then
         FromUTF8String = FromUTF8String & ChrW((b(0) And &HF) * CLng(&H1000) + (b(1) And &H3F) * CLng(&H40) + (b(2) And &H3F))
         i = i + 3
      ElseIf (b(0) And &HC0) = &HC0 Then
         FromUTF8String = FromUTF8String & ChrW((b(0) And &H1F) * &H40 + (b(1) And &H3F))
         i = i + 2
      Else
         FromUTF8String = FromUTF8String & Chr(b(0))
         i = i + 1
      End If
   Loop
End Function

