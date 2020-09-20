Attribute VB_Name = "modIPTC"
Option Explicit

    Public Type PICTDESC
        Size As Long
        Type As Long
        hHandle As Long
        hpal As Long
    End Type
    
    Public Declare Function timeGetTime Lib "winmm.dll" () As Long
    
    Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
    Public Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
    Public Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
    Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Public Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
    Public Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PICTDESC, riid As Any, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long
    
    Public Declare Function LoadImageW Lib "user32.dll" (ByVal hInst As Long, ByVal lpsz As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
    Public Const IMAGE_BITMAP As Long = 0
    Public Const LR_LOADFROMFILE As Long = &H10
    
    Public Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
    Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
    Public Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpFileSizeHigh As Long) As Long
    Public Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByRef lpOverlapped As Any) As Long
    Public Declare Function WriteFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByRef lpOverlapped As Any) As Long
    Private Declare Function GetFileAttributesW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long
    Private Declare Function SetFileAttributesW Lib "kernel32.dll" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long

    Public IPTCItemsDelimiter As String
    
    Public Type IPTCFields
        'alle bekannten 29 IPTC-Felder nach IPTC Specification IIMV4.1 (nicht IPTC Core XML)
        ObjectName As String
        Urgency As String
        Category As String
        SpecialInstructions As String
        DateCreated As String
        Byline As String
        BylineTitle As String
        City As String
        ProvinceState As String
        Country As String
        OriginalTransmissionReference As String
        Headline As String
        Credits As String
        Source As String
        Caption As String
        CaptionWriter As String
        TimeCreated As String
        Copyright As String
        EditStatus As String
        JobId As String
        ReleaseDate As String
        ReleaseTime As String
        OriginatingProgram As String
        ProgramVersion As String
        SubLocation As String
        LocationCode As String   'Location code
        Objectcycle As String
        SupplementalCategories As String
        Keywords As String
    End Type
    
    Dim SegMark As String
    Public iptc As IPTCFields

Public Function IPTCFromImage(FileName As String) As Boolean
    Dim hHandle As Long
    Dim imageData() As Byte
    Dim bytesRead As Long
    Dim IPTCHeader As String
    Dim IPTCPhotoshop As String
    Dim pos As Long
    Dim pos1 As Long
    Dim startHeader As Long
    Dim startPhotoshop As Long
    Dim strTemp As String
    Dim lngLen1 As Long
    Dim strIPTC As String
    Dim rc As Boolean
    Dim strImageString As String
    
    startHeader = 1
    startPhotoshop = 1
    IPTCHeader = Chr(&HFF&) & Chr(&HED&)                                'X'FFED'
    IPTCPhotoshop = Chr(&H50&) & Chr(&H68&) & Chr(&H6F&) & Chr(&H74&) & Chr(&H6F&) & Chr(&H73&)                         'Photoshop 3.0
    IPTCPhotoshop = IPTCPhotoshop & Chr(&H68&) & Chr(&H6F&) & Chr(&H70&) & Chr(&H20&) & Chr(&H33&) & Chr(&H2E&) & Chr(&H30&)
    hHandle = GetFileHandle(FileName, True)                             'true=read     'versteht unicode filename
    If hHandle <> INVALID_HANDLE_VALUE Then
        If hHandle Then
            bytesRead = GetFileSize(hHandle, ByVal 0&)
            If bytesRead Then
                ReDim imageData(0 To bytesRead - 1)
                ReadFile hHandle, imageData(0), bytesRead, bytesRead, ByVal 0&
                If bytesRead > UBound(imageData) Then
                    strImageString = ""
                    strImageString = StrConv(imageData, vbUnicode)
                End If
            End If
            CloseHandle hHandle
        End If
    End If
    Do Until startHeader >= bytesRead * 2 Or startPhotoshop >= bytesRead + 2
        pos = InStrB(startHeader, strImageString, IPTCHeader, vbTextCompare)                        'Header = X'FFED'  'Gerbing 22.05.2015
        pos1 = InStrB(startPhotoshop, strImageString, IPTCPhotoshop, vbTextCompare)
        If pos = 0 Or pos1 = 0 Then                                                                 'Gerbing 01.06.2015
            'Header oder Photoshop 3.0 nicht gefunden
            IPTCFromImage = False
            Exit Function
        Else
            'Header oder Photoshop 3.0 gefunden
            If pos1 = pos + 8 Then
                'sowohl Header wie Photoshop 3.0 gefunden im richtigen Abstand
                lngLen1 = LängeAusrechnen(strImageString, pos \ 2 + 3)                              'Gerbing 22.05.2015
                strIPTC = Mid(strImageString, pos \ 2, lngLen1 + 3)
                'strIPTC zerlegen in seine Einzelsegmente
                rc = VorhandeneEinzelsegmenteSuchen(strIPTC)
                IPTCFromImage = rc
                startHeader = pos + 1
                startPhotoshop = pos1 + 1
            Else
                'sowohl Header wie Photoshop 3.0 gefunden aber im falschen Abstand                  'Gerbing 01.06.2015
                If pos1 < pos Then
                    startPhotoshop = pos + 1
                Else
                    startHeader = pos1 - 8
                End If
            End If
        End If
    Loop
End Function

Private Function LängeAusrechnen(ss, pos) As Long
    'es sind generell 2 Bytes lange Felder
    Dim strLen1 As String           'Byte1
    Dim strlen2 As String           'Byte2
    Dim lngLen1 As Long
    Dim lnglen2 As Long
    
    strLen1 = Mid(ss, pos, 1)
    strlen2 = Mid(ss, pos + 1, 1)
    lngLen1 = Asc(strLen1)
    lnglen2 = Asc(strlen2)
    lngLen1 = lngLen1 * 256         'das erste Byte mit 256 multiplizieren
    lngLen1 = lngLen1 + lnglen2
    LängeAusrechnen = lngLen1
End Function

Private Function VorhandeneEinzelsegmenteSuchen(strIPTC As String) As Boolean
    'hier steht fest, dass es ein IPTC-Segment gibt
    'strIPTC beginnt mit dem IPTCHeader (X'FFED')
    'SegMark = Chr(&H1C&) & Chr(&H2&)                   X'1C02'
    'Segmentkennzeichen 1 Byte                          X'78'
    'Länge des Segmentes zb 32 Bytes                    X'0020'
    'Der Aufbau des IPTC-Segments ist in 'IPTC profile extracted.txt' beschrieben

    Dim pos As Long
    Dim start As Long
    Dim SegType As String
    Dim lngSeg As Long
    Dim SegMark As String
    Dim Datei As CEncodedFile

    SegMark = Chr(&H1C&) & Chr(&H2&)                   'X'1C02'
    start = 1
    iptc.SupplementalCategories = ""
    iptc.Keywords = ""
    Do Until start > Len(strIPTC)
        'SegmentMarker suchen
        pos = InStr(start, strIPTC, SegMark, vbBinaryCompare)
        If pos = 0 Then
            VorhandeneEinzelsegmenteSuchen = False
            Exit Function
        End If
        SegType = Mid(strIPTC, pos + 2, 1)
        Select Case SegType
            Case Chr(&H78&)                                 '78=Caption
                lngSeg = LängeAusrechnen(strIPTC, pos + 3)
                iptc.Caption = Mid(strIPTC, pos + 5, lngSeg)
                start = pos + lngSeg + 5
'                Set Datei = New CEncodedFile
'                Datei.Load (iptc.Caption)
'                iptc.Caption = Mid(Datei.Text, 1, Len(Datei.Text) \ 2)
                iptc.Caption = FromUTF8String(Mid(iptc.Caption, 1))                             'Gerbing 16.11.2015
                Set Datei = Nothing
            Case Chr(&H41&)                                 '41=Originating Program
                lngSeg = LängeAusrechnen(strIPTC, pos + 3)
                iptc.OriginatingProgram = Mid(strIPTC, pos + 5, lngSeg)
                start = pos + lngSeg + 5
'                Set Datei = New CEncodedFile
'                Datei.Load (iptc.OriginatingProgram)
'                iptc.OriginatingProgram = Mid(Datei.Text, 1, Len(Datei.Text) \ 2)
                iptc.OriginatingProgram = FromUTF8String(Mid(iptc.OriginatingProgram, 1))       'Gerbing 16.11.2015
                Set Datei = Nothing
            Case Chr(&H5&)                                  '05=ObjectName
                lngSeg = LängeAusrechnen(strIPTC, pos + 3)
                iptc.ObjectName = Mid(strIPTC, pos + 5, lngSeg)
                start = pos + lngSeg + 5
'                Set Datei = New CEncodedFile
'                Datei.Load (iptc.ObjectName)
'                iptc.ObjectName = Mid(Datei.Text, 1, Len(Datei.Text) \ 2)
                iptc.ObjectName = FromUTF8String(Mid(iptc.ObjectName, 1))                       'Gerbing 16.11.2015
                Set Datei = Nothing
            Case Chr(&HA&)                                  '0A=Urgency
                lngSeg = LängeAusrechnen(strIPTC, pos + 3)
                iptc.Urgency = Mid(strIPTC, pos + 5, lngSeg)
                start = pos + lngSeg + 5
'                Set Datei = New CEncodedFile
'                Datei.Load (iptc.Urgency)
'                iptc.Urgency = Mid(Datei.Text, 1, Len(Datei.Text) \ 2)
                iptc.Urgency = FromUTF8String(Mid(iptc.Urgency, 1))                             'Gerbing 16.11.2015
                Set Datei = Nothing
            Case Chr(&HF&)                                  '0F=Category
                lngSeg = LängeAusrechnen(strIPTC, pos + 3)
                iptc.Category = Mid(strIPTC, pos + 5, lngSeg)
                start = pos + lngSeg + 5
'                Set Datei = New CEncodedFile
'                Datei.Load (iptc.Category)
'                iptc.Category = Mid(Datei.Text, 1, Len(Datei.Text) \ 2)
                iptc.Category = FromUTF8String(Mid(iptc.Category, 1))                           'Gerbing 16.11.2015
                Set Datei = Nothing
            Case Chr(&H28&)                                 '28=SpecialInstructions
                lngSeg = LängeAusrechnen(strIPTC, pos + 3)
                iptc.SpecialInstructions = Mid(strIPTC, pos + 5, lngSeg)
                start = pos + lngSeg + 5
'                Set Datei = New CEncodedFile
'                Datei.Load (iptc.SpecialInstructions)
'                iptc.SpecialInstructions = Mid(Datei.Text, 1, Len(Datei.Text) \ 2)
                iptc.SpecialInstructions = FromUTF8String(Mid(iptc.SpecialInstructions, 1))     'Gerbing 16.11.2015
                Set Datei = Nothing
            Case Chr(&H37&)                                 '37=DateCreated
                lngSeg = LängeAusrechnen(strIPTC, pos + 3)
                iptc.DateCreated = Mid(strIPTC, pos + 5, lngSeg)
                start = pos + lngSeg + 5
'                Set Datei = New CEncodedFile
'                Datei.Load (iptc.DateCreated)
'                iptc.DateCreated = Mid(Datei.Text, 1, Len(Datei.Text) \ 2)
                iptc.DateCreated = FromUTF8String(Mid(iptc.DateCreated, 1))                     'Gerbing 16.11.2015
                Set Datei = Nothing
            Case Chr(&H50&)                                 '50=Byline
                lngSeg = LängeAusrechnen(strIPTC, pos + 3)
                iptc.Byline = Mid(strIPTC, pos + 5, lngSeg)
                start = pos + lngSeg + 5
'                Set Datei = New CEncodedFile
'                Datei.Load (iptc.Byline)
'                iptc.Byline = Mid(Datei.Text, 1, Len(Datei.Text) \ 2)
                iptc.Byline = FromUTF8String(Mid(iptc.Byline, 1))                               'Gerbing 16.11.2015
                Set Datei = Nothing
            Case Chr(&H55&)                                 '55=BylineTitle
                lngSeg = LängeAusrechnen(strIPTC, pos + 3)
                iptc.BylineTitle = Mid(strIPTC, pos + 5, lngSeg)
                start = pos + lngSeg + 5
'                Set Datei = New CEncodedFile
'                Datei.Load (iptc.BylineTitle)
'                iptc.BylineTitle = Mid(Datei.Text, 1, Len(Datei.Text) \ 2)
                iptc.BylineTitle = FromUTF8String(Mid(iptc.BylineTitle, 1))                     'Gerbing 16.11.2015
                Set Datei = Nothing
            Case Chr(&H5A&)                                 '5A=City
                lngSeg = LängeAusrechnen(strIPTC, pos + 3)
                iptc.City = Mid(strIPTC, pos + 5, lngSeg)
                start = pos + lngSeg + 5
'                Set Datei = New CEncodedFile
'                Datei.Load (iptc.City)
'                iptc.City = Mid(Datei.Text, 1, Len(Datei.Text) \ 2)
                iptc.City = FromUTF8String(Mid(iptc.City, 1))                                   'Gerbing 16.11.2015
                Set Datei = Nothing
            Case Chr(&H5F&)                                 '5F=State
                lngSeg = LängeAusrechnen(strIPTC, pos + 3)
                iptc.ProvinceState = Mid(strIPTC, pos + 5, lngSeg)
                start = pos + lngSeg + 5
'                Set Datei = New CEncodedFile
'                Datei.Load (iptc.ProvinceState)
'                iptc.ProvinceState = Mid(Datei.Text, 1, Len(Datei.Text) \ 2)
                iptc.ProvinceState = FromUTF8String(Mid(iptc.ProvinceState, 1))                 'Gerbing 16.11.2015
                Set Datei = Nothing
            Case Chr(&H65&)                                 '65=Country
                lngSeg = LängeAusrechnen(strIPTC, pos + 3)
                iptc.Country = Mid(strIPTC, pos + 5, lngSeg)
                start = pos + lngSeg + 5
'                Set Datei = New CEncodedFile
'                Datei.Load (iptc.Country)
'                iptc.Country = Mid(Datei.Text, 1, Len(Datei.Text) \ 2)
                iptc.Country = FromUTF8String(Mid(iptc.Country, 1))                                 'Gerbing 16.11.2015
                Set Datei = Nothing
            Case Chr(&H67&)                                 '67=OriginalTransmissionReference
                lngSeg = LängeAusrechnen(strIPTC, pos + 3)
                iptc.OriginalTransmissionReference = Mid(strIPTC, pos + 5, lngSeg)
                start = pos + lngSeg + 5
'                Set Datei = New CEncodedFile
'                Datei.Load (iptc.OriginalTransmissionReference)
'                iptc.OriginalTransmissionReference = Mid(Datei.Text, 1, Len(Datei.Text) \ 2)
                iptc.OriginalTransmissionReference = FromUTF8String(Mid(iptc.OriginalTransmissionReference, 1)) 'Gerbing 16.11.2015
                Set Datei = Nothing
            Case Chr(&H69&)                                 '69=Headline
                lngSeg = LängeAusrechnen(strIPTC, pos + 3)
                iptc.Headline = Mid(strIPTC, pos + 5, lngSeg)
                start = pos + lngSeg + 5
'                Set Datei = New CEncodedFile
'                Datei.Load (iptc.Headline)
'                iptc.Headline = Mid(Datei.Text, 1, Len(Datei.Text) \ 2)
                iptc.Headline = FromUTF8String(Mid(iptc.Headline, 1))                               'Gerbing 16.11.2015
                Set Datei = Nothing
            Case Chr(&H6E&)                                 '6E=Credits
                lngSeg = LängeAusrechnen(strIPTC, pos + 3)
                iptc.Credits = Mid(strIPTC, pos + 5, lngSeg)
                start = pos + lngSeg + 5
'                Set Datei = New CEncodedFile
'                Datei.Load (iptc.Credits)
'                iptc.Credits = Mid(Datei.Text, 1, Len(Datei.Text) \ 2)
                iptc.Credits = FromUTF8String(Mid(iptc.Credits, 1))                                 'Gerbing 16.11.2015
                Set Datei = Nothing
            Case Chr(&H73&)                                 '73=Source
                lngSeg = LängeAusrechnen(strIPTC, pos + 3)
                iptc.Source = Mid(strIPTC, pos + 5, lngSeg)
                start = pos + lngSeg + 5
'                Set Datei = New CEncodedFile
'                Datei.Load (iptc.Source)
'                iptc.Source = Mid(Datei.Text, 1, Len(Datei.Text) \ 2)
                iptc.Source = FromUTF8String(Mid(iptc.Source, 1))                                   'Gerbing 16.11.2015
                Set Datei = Nothing
            Case Chr(&H7A&)                                 '7A=CaptionWriter
                lngSeg = LängeAusrechnen(strIPTC, pos + 3)
                iptc.CaptionWriter = Mid(strIPTC, pos + 5, lngSeg)
                start = pos + lngSeg + 5
'                Set Datei = New CEncodedFile
'                Datei.Load (iptc.CaptionWriter)
'                iptc.CaptionWriter = Mid(Datei.Text, 1, Len(Datei.Text) \ 2)
                iptc.CaptionWriter = FromUTF8String(Mid(iptc.CaptionWriter, 1))                     'Gerbing 16.11.2015
                Set Datei = Nothing
            Case Chr(&H3C&)                                 '3C=TimeCreated
                lngSeg = LängeAusrechnen(strIPTC, pos + 3)
                iptc.TimeCreated = Mid(strIPTC, pos + 5, lngSeg)
                start = pos + lngSeg + 5
'                Set Datei = New CEncodedFile
'                Datei.Load (iptc.TimeCreated)
'                iptc.TimeCreated = Mid(Datei.Text, 1, Len(Datei.Text) \ 2)
                iptc.TimeCreated = FromUTF8String(Mid(iptc.TimeCreated, 1))                         'Gerbing 16.11.2015
                Set Datei = Nothing
            Case Chr(&H74&)                                 '74=Copyright
                lngSeg = LängeAusrechnen(strIPTC, pos + 3)
                iptc.Copyright = Mid(strIPTC, pos + 5, lngSeg)
                start = pos + lngSeg + 5
'                Set Datei = New CEncodedFile
'                Datei.Load (iptc.Copyright)
'                iptc.Copyright = Mid(Datei.Text, 1, Len(Datei.Text) \ 2)
                iptc.Copyright = FromUTF8String(Mid(iptc.Copyright, 1))                             'Gerbing 16.11.2015
                Set Datei = Nothing
            Case Chr(&H7&)                                  '07=EditStatus
                lngSeg = LängeAusrechnen(strIPTC, pos + 3)
                iptc.EditStatus = Mid(strIPTC, pos + 5, lngSeg)
                start = pos + lngSeg + 5
'                Set Datei = New CEncodedFile
'                Datei.Load (iptc.EditStatus)
'                iptc.EditStatus = Mid(Datei.Text, 1, Len(Datei.Text) \ 2)
                iptc.EditStatus = FromUTF8String(Mid(iptc.EditStatus, 1))                           'Gerbing 16.11.2015
                Set Datei = Nothing
            Case Chr(&H16&)                                 '16=JobId
                lngSeg = LängeAusrechnen(strIPTC, pos + 3)
                iptc.JobId = Mid(strIPTC, pos + 5, lngSeg)
                start = pos + lngSeg + 5
'                Set Datei = New CEncodedFile
'                Datei.Load (iptc.JobId)
'                iptc.JobId = Mid(Datei.Text, 1, Len(Datei.Text) \ 2)
                iptc.JobId = FromUTF8String(Mid(iptc.JobId, 1))                                     'Gerbing 16.11.2015
                Set Datei = Nothing
            Case Chr(&H1E&)                                 '1E=ReleaseDate
                lngSeg = LängeAusrechnen(strIPTC, pos + 3)
                iptc.ReleaseDate = Mid(strIPTC, pos + 5, lngSeg)
                start = pos + lngSeg + 5
'                Set Datei = New CEncodedFile
'                Datei.Load (iptc.ReleaseDate)
'                iptc.ReleaseDate = Mid(Datei.Text, 1, Len(Datei.Text) \ 2)
                iptc.ReleaseDate = FromUTF8String(Mid(iptc.ReleaseDate, 1))                         'Gerbing 16.11.2015
                Set Datei = Nothing
            Case Chr(&H23&)                                 '23=ReleaseTime
                lngSeg = LängeAusrechnen(strIPTC, pos + 3)
                iptc.ReleaseTime = Mid(strIPTC, pos + 5, lngSeg)
                start = pos + lngSeg + 5
'                Set Datei = New CEncodedFile
'                Datei.Load (iptc.ReleaseTime)
'                iptc.ReleaseTime = Mid(Datei.Text, 1, Len(Datei.Text) \ 2)
                iptc.ReleaseTime = FromUTF8String(Mid(iptc.ReleaseTime, 1))                         'Gerbing 16.11.2015
                Set Datei = Nothing
            Case Chr(&H46&)                                 '46=ProgramVersion
                lngSeg = LängeAusrechnen(strIPTC, pos + 3)
                iptc.ProgramVersion = Mid(strIPTC, pos + 5, lngSeg)
                start = pos + lngSeg + 5
'                Set Datei = New CEncodedFile
'                Datei.Load (iptc.ProgramVersion)
'                iptc.ProgramVersion = Mid(Datei.Text, 1, Len(Datei.Text) \ 2)
                iptc.ProgramVersion = FromUTF8String(Mid(iptc.ProgramVersion, 1))                   'Gerbing 16.11.2015
                Set Datei = Nothing
            Case Chr(&H5C&)                                 '5C=Sublocation
                lngSeg = LängeAusrechnen(strIPTC, pos + 3)
                iptc.SubLocation = Mid(strIPTC, pos + 5, lngSeg)
                start = pos + lngSeg + 5
'                Set Datei = New CEncodedFile
'                Datei.Load (iptc.SubLocation)
'                iptc.SubLocation = Mid(Datei.Text, 1, Len(Datei.Text) \ 2)
                iptc.SubLocation = FromUTF8String(Mid(iptc.SubLocation, 1))                         'Gerbing 16.11.2015
                Set Datei = Nothing
            Case Chr(&H64&)                                 '64=LocationCode
                lngSeg = LängeAusrechnen(strIPTC, pos + 3)
                iptc.LocationCode = Mid(strIPTC, pos + 5, lngSeg)
                start = pos + lngSeg + 5
'                Set Datei = New CEncodedFile
'                Datei.Load (iptc.LocationCode)
'                iptc.LocationCode = Mid(Datei.Text, 1, Len(Datei.Text) \ 2)
                iptc.LocationCode = FromUTF8String(Mid(iptc.LocationCode, 1))                       'Gerbing 16.11.2015
                Set Datei = Nothing
            Case Chr(&H4B&)                                 '4B=Objectcycle
                lngSeg = LängeAusrechnen(strIPTC, pos + 3)
                iptc.Objectcycle = Mid(strIPTC, pos + 5, lngSeg)
                start = pos + lngSeg + 5
'                Set Datei = New CEncodedFile
'                Datei.Load (iptc.Objectcycle)
'                iptc.Objectcycle = Mid(Datei.Text, 1, Len(Datei.Text) \ 2)
                iptc.Objectcycle = FromUTF8String(Mid(iptc.Objectcycle, 1))                         'Gerbing 16.11.2015
                Set Datei = Nothing
            Case Chr(&H14&)                                 '14=Supplemental Categories
                'Supplemental Categories können mehrfach auftreten
                lngSeg = LängeAusrechnen(strIPTC, pos + 3)
                If iptc.SupplementalCategories = "" Then
                    iptc.SupplementalCategories = Mid(strIPTC, pos + 5, lngSeg)
                Else
                    iptc.SupplementalCategories = iptc.SupplementalCategories & Mid(strIPTC, pos + 5, lngSeg)
                End If
                start = pos + lngSeg + 5
'                Set Datei = New CEncodedFile
'                Datei.Load (iptc.SupplementalCategories)
'                iptc.SupplementalCategories = Mid(Datei.Text, 1, Len(Datei.Text) \ 2)
                iptc.SupplementalCategories = FromUTF8String(Mid(iptc.SupplementalCategories, 1)) 'Gerbing 16.11.2015
                Set Datei = Nothing
            Case Chr(&H19&)                                 '19=Keywords
                'Keywords können mehrfach auftreten
                lngSeg = LängeAusrechnen(strIPTC, pos + 3)
                If iptc.Keywords = "" Then
                    iptc.Keywords = Mid(strIPTC, pos + 5, lngSeg)
                Else
                    iptc.Keywords = iptc.Keywords & Mid(strIPTC, pos + 5, lngSeg)
                End If
                start = pos + lngSeg + 5
'                Set Datei = New CEncodedFile
'                Datei.Load (iptc.Keywords)
'                iptc.Keywords = Mid(Datei.Text, 1, Len(Datei.Text) \ 2)
                iptc.Keywords = FromUTF8String(Mid(iptc.Keywords, 1))                         'Gerbing 16.11.2015
                Set Datei = Nothing
            Case Else
                'unbekannte Segenttypen werden nicht gerettet
                lngSeg = LängeAusrechnen(strIPTC, pos + 3)
                start = pos + lngSeg + 5
        End Select
    Loop
    VorhandeneEinzelsegmenteSuchen = True
End Function

Public Function LoadPictureW(FileName As String, meForm As Form)
    Dim hImage As Long
    Dim tmpPic As StdPicture
    Dim hHandle As Long
    Dim imageData() As Byte
    Dim bytesRead As Long

    If StrComp(Right(FileName, 3), "bmp", vbTextCompare) = 0 Then
        hImage = LoadImageW(0&, StrPtr(FileName), IMAGE_BITMAP, 0&, 0&, LR_LOADFROMFILE)
        If hImage Then Set tmpPic = HandleToStdPicture(hImage, vbPicTypeBitmap)
    Else ' loaded a gif, jpg or possibly something else?
        hHandle = GetFileHandle(FileName, True)                 'true=read
        If hHandle <> INVALID_HANDLE_VALUE Then
            If hHandle Then
                bytesRead = GetFileSize(hHandle, ByVal 0&)
                If bytesRead Then
                    ReDim imageData(0 To bytesRead - 1)
                    ReadFile hHandle, imageData(0), bytesRead, bytesRead, ByVal 0&
                    If bytesRead > UBound(imageData) Then
                        Set tmpPic = ArrayToPicture(imageData(), 0, bytesRead)
                    End If
                End If
                CloseHandle hHandle
            End If
        End If
    End If
    If tmpPic Is Nothing Then
        'MsgBox "Vorschaubild nicht möglich", vbOKOnly + vbExclamation
        MsgBox LoadResString(2462 + Sprache), vbOKOnly + vbExclamation
    Else
        Set meForm.Picture1.Picture = tmpPic
    End If
End Function

Private Function HandleToStdPicture(ByVal hImage As Long, ByVal imgType As Long) As IPicture
    ' function creates a stdPicture object from an image handle (bitmap or icon)
    Dim lpPictDesc As PICTDESC, aGUID(0 To 3) As Long
    With lpPictDesc
        .Size = Len(lpPictDesc)
        .Type = imgType
        .hHandle = hImage
        .hpal = 0
    End With
    ' IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    aGUID(0) = &H7BF80980
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    ' create stdPicture
    Call OleCreatePictureIndirect(lpPictDesc, aGUID(0), True, HandleToStdPicture)

End Function

Private Function ArrayToPicture(inArray() As Byte, Offset As Long, Size As Long) As IPicture
    ' function creates a stdPicture from the passed array
    ' Note: The array was already validated as not empty when calling class' LoadStream was called
    Dim o_hMem  As Long
    Dim o_lpMem  As Long
    Dim aGUID(0 To 3) As Long
    Dim IIStream As IUnknown
    
    aGUID(0) = &H7BF80980    ' GUID for stdPicture
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    
    o_hMem = GlobalAlloc(&H2&, Size)
    If Not o_hMem = 0& Then
        o_lpMem = GlobalLock(o_hMem)
        If Not o_lpMem = 0& Then
            CopyMemory ByVal o_lpMem, inArray(Offset), Size
            Call GlobalUnlock(o_hMem)
            If CreateStreamOnHGlobal(o_hMem, 1&, IIStream) = 0& Then
                  Call OleLoadPicture(ByVal ObjPtr(IIStream), 0&, 0&, aGUID(0), ArrayToPicture)
            End If
        End If
    End If
End Function

Public Function LeseIPTC(Fotodatei As String, LstU As Control, Ausgeben As Boolean) As Boolean
    Dim rc As Boolean
    Dim start As Long
    Const Standardlänge As Long = 70
    Dim Länge As Long
    Dim strTemp As String

    'rc = false = 0 wenn kein IPTC Feld gefunden
    'rc = True = -1 wenn mindestens 1 IPTC Feld gefunden
    
    Call IPTCFelderLöschen
    IPTCItemsDelimiter = ";"
    rc = IPTCFromImage(Fotodatei)
    DoEvents                                                                                                'Gerbing 22.05.2015
    
    If iptc.OriginatingProgram = "" And iptc.ObjectName = "" And iptc.Byline = "" And iptc.BylineTitle = "" _
        And iptc.Caption = "" And iptc.CaptionWriter = "" And iptc.Copyright = "" And _
        iptc.SpecialInstructions = "" And iptc.Urgency = "" And iptc.DateCreated = "" And _
        iptc.TimeCreated = "" And iptc.City = "" And iptc.ProvinceState = "" And iptc.Country = "" And _
        iptc.Credits = "" And iptc.Source = "" And iptc.Headline = "" And iptc.OriginalTransmissionReference = "" _
        And iptc.Category = "" And iptc.SupplementalCategories = "" And iptc.Keywords = "" _
 _
        And iptc.ReleaseDate = "" And iptc.ReleaseTime = "" And iptc.Objectcycle = "" And iptc.LocationCode = "" _
        And iptc.SubLocation = "" And iptc.ProgramVersion = "" And iptc.EditStatus = "" And iptc.JobId = "" Then
        LeseIPTC = False
        Exit Function
    Else
        If Ausgeben = False Then
            LeseIPTC = True
            Exit Function
        End If
        '------------------------------------------------------------------------------------------------------------
        'hierher kommt es nur, wenn Ausgeben = True
        If iptc.OriginatingProgram <> "" Then
            LstU.ListItems.Add "OriginatingProgram" & " - " & iptc.OriginatingProgram
        End If
        If iptc.ObjectName <> "" Then
            LstU.ListItems.Add "ObjectName" & " - " & iptc.ObjectName
        End If
        If iptc.Byline <> "" Then
            LstU.ListItems.Add "Byline" & " - " & iptc.Byline
        End If
        If iptc.BylineTitle <> "" Then
            LstU.ListItems.Add "Byline title" & " - " & iptc.BylineTitle
        End If
        If iptc.Caption <> "" Then
            start = 1
            Länge = Standardlänge
            If Len(iptc.Caption) < Länge Then
                Länge = Len(iptc.Caption)
            End If
            strTemp = ""
            Do Until start > (Len(iptc.Caption))
                strTemp = Mid(iptc.Caption, start, Länge)
                start = start + Länge
                LstU.ListItems.Add "Caption" & " - " & strTemp
            Loop
        End If
        If iptc.CaptionWriter <> "" Then
            LstU.ListItems.Add "Caption writer" & " - " & iptc.CaptionWriter
        End If
        If iptc.Copyright <> "" Then
            LstU.ListItems.Add "Copyright notice" & " - " & iptc.Copyright
        End If
        If iptc.SpecialInstructions <> "" Then
            start = 1
            Länge = Standardlänge
            If Len(iptc.SpecialInstructions) < Länge Then
                Länge = Len(iptc.SpecialInstructions)
            End If
            strTemp = ""
            Do Until start > (Len(iptc.SpecialInstructions))
                strTemp = Mid(iptc.SpecialInstructions, start, Länge)
                start = start + Länge
                LstU.ListItems.Add "Special Instructions" & " - " & strTemp
            Loop
        End If
        If iptc.Urgency <> "" Then
            LstU.ListItems.Add "Urgency" & " - " & iptc.Urgency
        End If
        If iptc.DateCreated <> "" Then
            'LstU.ListItems.Add "Date created" & " - " & Mid(IPTC.DateCreated, 5, 2) & "/" & Mid(IPTC.DateCreated, 7, 2) & "/" & Mid(IPTC.DateCreated, 1, 4)
            LstU.ListItems.Add "Date created" & " - " & iptc.DateCreated
        End If
        If iptc.TimeCreated <> "" Then
            'LstU.ListItems.Add "Time created" & " - " & Mid(IPTC.TimeCreated, 1, 2) & ":" & Mid(IPTC.TimeCreated, 3, 2) & ":" & Mid(IPTC.TimeCreated, 5, 2) & " GMT" & Mid(IPTC.TimeCreated, 7)
            LstU.ListItems.Add "Time created" & " - " & iptc.TimeCreated
        End If
        If iptc.City <> "" Then                                                         'Gerbing 04.03.2013
            'MessageBoxW 0, StrPtr(Datei.Text), StrPtr(Datei.Text), MB_ICONINFORMATION Or MB_TASKMODAL
            LstU.ListItems.Add "City" & " - " & iptc.City
        End If
        If iptc.ProvinceState <> "" Then
            LstU.ListItems.Add "Province/State" & " - " & iptc.ProvinceState
        End If
        If iptc.Country <> "" Then
            LstU.ListItems.Add "Country" & " - " & iptc.Country
        End If
        If iptc.Credits <> "" Then
            LstU.ListItems.Add "Credit" & " - " & iptc.Credits
        End If
        If iptc.Source <> "" Then
            LstU.ListItems.Add "Source" & " - " & iptc.Source
        End If
        If iptc.Headline <> "" Then
            start = 1
            Länge = Standardlänge
                    If Len(iptc.Headline) < Länge Then
                Länge = Len(iptc.Headline)
            End If
            strTemp = ""
            Do Until start > (Len(iptc.Headline))
                strTemp = Mid(iptc.Headline, start, Länge)
                start = start + Länge
                LstU.ListItems.Add "Headline" & " - " & strTemp
            Loop
        End If
        If iptc.OriginalTransmissionReference <> "" Then
            LstU.ListItems.Add "Transmission reference" & " - " & iptc.OriginalTransmissionReference
        End If
        If iptc.Category <> "" Then
            LstU.ListItems.Add "Category" & " - " & iptc.Category
        End If
        If iptc.SupplementalCategories <> "" Then
            start = 1
            Länge = Standardlänge
            If Len(iptc.SupplementalCategories) < Länge Then
                Länge = Len(iptc.SupplementalCategories)
            End If
            strTemp = ""
            Do Until start > (Len(iptc.SupplementalCategories))
                strTemp = Mid(iptc.SupplementalCategories, start, Länge)
                start = start + Länge
                LstU.ListItems.Add "SupplementalCategories" & " - " & strTemp
            Loop
        End If
        If iptc.Keywords <> "" Then
            start = 1
            Länge = Standardlänge
            If Len(iptc.Keywords) < Länge Then
                Länge = Len(iptc.Keywords)
            End If
            strTemp = ""
            Do Until start > (Len(iptc.Keywords))
                strTemp = Mid(iptc.Keywords, start, Länge)
                start = start + Länge
                LstU.ListItems.Add "Keywords" & " - " & strTemp
            Loop
        End If
        '-------------------------------------------------------
        'jetzt die restlichen 8 IPTC-Felder
        '-------------------------------------------------------
        If iptc.ReleaseDate <> "" Then
            LstU.ListItems.Add "ReleaseDate" & " - " & iptc.ReleaseDate
        End If
        If iptc.ReleaseTime <> "" Then
            LstU.ListItems.Add "ReleaseTime" & " - " & iptc.ReleaseTime
        End If
        If iptc.Objectcycle <> "" Then
            LstU.ListItems.Add "Objectcycle" & " - " & iptc.Objectcycle
        End If
        If iptc.LocationCode <> "" Then
            LstU.ListItems.Add "LocationCode" & " - " & iptc.LocationCode
        End If
        If iptc.SubLocation <> "" Then
            LstU.ListItems.Add "SubLocation" & " - " & iptc.SubLocation
        End If
        If iptc.ProgramVersion <> "" Then
            LstU.ListItems.Add "ProgramVersion" & " - " & iptc.ProgramVersion
        End If
        If iptc.EditStatus <> "" Then
            LstU.ListItems.Add "EditStatus" & " - " & iptc.EditStatus
        End If
        If iptc.JobId <> "" Then
            LstU.ListItems.Add "JobId" & " - " & iptc.JobId
        End If
        LeseIPTC = True
    End If
End Function

Public Function GetFileHandle(FileName As String, bRead As Boolean) As Long

    ' Function uses APIs to read/create files with unicode support

    Const GENERIC_READ As Long = &H80000000
    Const OPEN_EXISTING = &H3
    Const FILE_SHARE_READ = &H1
    Const GENERIC_WRITE As Long = &H40000000
    Const FILE_SHARE_WRITE As Long = &H2
    Const CREATE_ALWAYS As Long = 2
    Const FILE_ATTRIBUTE_ARCHIVE As Long = &H20
    Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
    Const FILE_ATTRIBUTE_READONLY As Long = &H1
    Const FILE_ATTRIBUTE_SYSTEM As Long = &H4
    Const FILE_ATTRIBUTE_NORMAL = &H80&
    
    Dim flags As Long, Access As Long
    Dim Disposition As Long, Share As Long, hFile As Long

    If bRead Then
        Access = GENERIC_READ
        Share = FILE_SHARE_READ
        Disposition = OPEN_EXISTING
        flags = FILE_ATTRIBUTE_ARCHIVE Or FILE_ATTRIBUTE_HIDDEN Or FILE_ATTRIBUTE_NORMAL _
                Or FILE_ATTRIBUTE_READONLY Or FILE_ATTRIBUTE_SYSTEM
    Else
        Access = GENERIC_READ Or GENERIC_WRITE
        Share = 0&
        flags = GetFileAttributesW(StrPtr(FileName))
        If (flags And FILE_ATTRIBUTE_READONLY) Then
            flags = FILE_ATTRIBUTE_NORMAL
            SetFileAttributesW StrPtr(FileName), flags
        End If
        If flags < 0& Then flags = FILE_ATTRIBUTE_NORMAL
        ' CREATE_ALWAYS will delete previous file if necessary
        Disposition = CREATE_ALWAYS
    End If
    hFile = CreateFileW(StrPtr(FileName), Access, Share, ByVal 0&, Disposition, flags, 0&)
    If hFile = 0& Then hFile = INVALID_HANDLE_VALUE
    GetFileHandle = hFile
End Function


Private Sub IPTCFelderLöschen()
    iptc.OriginatingProgram = ""
    iptc.ObjectName = ""
    iptc.Byline = ""
    iptc.BylineTitle = ""
    iptc.Caption = ""
    iptc.CaptionWriter = ""
    iptc.Copyright = ""
    iptc.SpecialInstructions = ""
    iptc.Urgency = ""
    iptc.DateCreated = ""
    iptc.TimeCreated = ""
    iptc.City = ""
    iptc.ProvinceState = ""
    iptc.Country = ""
    iptc.Credits = ""
    iptc.Source = ""
    iptc.Headline = ""
    iptc.OriginalTransmissionReference = ""
    iptc.Category = ""
    iptc.SupplementalCategories = ""
    iptc.Keywords = ""
    iptc.ReleaseDate = ""
    iptc.ReleaseTime = ""
    iptc.Objectcycle = ""
    iptc.LocationCode = ""
    iptc.SubLocation = ""
    iptc.ProgramVersion = ""
    iptc.EditStatus = ""
    iptc.JobId = ""
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

