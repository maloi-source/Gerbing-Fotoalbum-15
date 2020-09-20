VERSION 5.00
Object = "{A8F9B8E7-E699-4FCE-A647-72C877F8E632}#1.8#0"; "editctlsu.ocx"
Begin VB.Form frmFeldAktualisierung 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Feld-Aktualisierung durch Import-Wiederholung"
   ClientHeight    =   2196
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   10512
   Icon            =   "frmFeldAktualisierung.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2196
   ScaleWidth      =   10512
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton btnStart 
      Caption         =   "Start"
      Height          =   492
      Left            =   3120
      TabIndex        =   3
      Top             =   480
      Width           =   1932
   End
   Begin VB.CheckBox chkEXIFDateTimeOriginal 
      BackColor       =   &H00C0C0C0&
      Caption         =   "EXIFDateTimeOriginal"
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2772
   End
   Begin VB.CheckBox chkGPSLongitude 
      BackColor       =   &H00C0C0C0&
      Caption         =   "GPSLongitude"
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2772
   End
   Begin VB.CheckBox chkGPSLatitude 
      BackColor       =   &H00C0C0C0&
      Caption         =   "GPSLatitude"
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2772
   End
   Begin EditCtlsLibUCtl.TextBox txtArbeitsfortschritt 
      Height          =   492
      Left            =   3120
      TabIndex        =   5
      Top             =   1560
      Width           =   7332
      _cx             =   12933
      _cy             =   868
      AcceptNumbersOnly=   0   'False
      AcceptTabKey    =   0   'False
      AllowDragDrop   =   -1  'True
      AlwaysShowSelection=   0   'False
      Appearance      =   1
      AutoScrolling   =   2
      BackColor       =   -2147483643
      BorderStyle     =   0
      CancelIMECompositionOnSetFocus=   0   'False
      CharacterConversion=   0
      CompleteIMECompositionOnKillFocus=   0   'False
      DisabledBackColor=   -1
      DisabledEvents  =   3075
      DisabledForeColor=   -1
      DisplayCueBannerOnFocus=   0   'False
      DontRedraw      =   0   'False
      DoOEMConversion =   0   'False
      DragScrollTimeBase=   -1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      FormattingRectangleHeight=   0
      FormattingRectangleLeft=   0
      FormattingRectangleTop=   0
      FormattingRectangleWidth=   0
      HAlignment      =   0
      HoverTime       =   -1
      IMEMode         =   -1
      InsertMarkColor =   0
      InsertSoftLineBreaks=   0   'False
      LeftMargin      =   -1
      MaxTextLength   =   -1
      Modified        =   0   'False
      MousePointer    =   0
      MultiLine       =   0   'False
      OLEDragImageStyle=   0
      PasswordChar    =   0
      ProcessContextMenuKeys=   -1  'True
      ReadOnly        =   0   'False
      RegisterForOLEDragDrop=   0   'False
      RightMargin     =   -1
      RightToLeft     =   0
      ScrollBars      =   0
      SelectedTextMousePointer=   0
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      UseCustomFormattingRectangle=   0   'False
      UsePasswordChar =   0   'False
      UseSystemFont   =   0   'False
      CueBanner       =   "frmFeldAktualisierung.frx":038A
      Text            =   "frmFeldAktualisierung.frx":03AA
   End
   Begin VB.Label lblArbeitsfortschritt 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Arbeitsfortschritt:"
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   2892
   End
End
Attribute VB_Name = "frmFeldAktualisierung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim blnMediaInfoInitialized As Boolean                                          'Gerbing 18.11.2019
    Dim Handle As Long                                                              'Gerbing 18.11.2019

Private Sub btnStart_Click()
    Dim SQL As String
    Dim tempDateiname As String
    Dim EXIFInfo As String
    Dim strEXIFDateTimeOriginal As String
    Dim strGPSLatitude As String
    Dim strGPSLongitude As String
    Dim DateinamenErweiterung As String                                                 'Gerbing 18.11.2019
    Dim Msg As String
    Dim Pos As Long
    Dim pos1 As Long
    Dim pos2 As Long
    Dim antwort As Long                                                                 'Gerbing 18.11.2019
    
    'Für alle Fotos im Suchergebnis soll GPSLatitude, GPSLongitude, EXIFDateTime aktualisiert werden
    'Bei mp4 und mov Videos brauche ich die MediaInfo.DLL                               'Gerbing 18.11.2019
    Screen.MousePointer = vbHourglass
    If gblnSchreibgeschützt = False Then
        SQL = frmGridAndThumb.rsDataGrid.Source
        With rstsql
            .Source = SQL
            .ActiveConnection = DBado                                                   'Gerbing 23.11.2017
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
        Do Until rstsql.EOF
            tempDateiname = Replace(rstsql.Fields(LoadResString(1028 + Sprache)), "+:\", gstrFotosMdbLocation & "\")
            DateinamenErweiterung = Right(tempDateiname, 3)                             'Gerbing 18.11.2019
            DateinamenErweiterung = UCase(DateinamenErweiterung)
            Form1.EXF.ImageFile = tempDateiname 'set the image file property, read metainfo, parse metainfo
            '
            'EXF.ListInfo ist ein String mit vbCrLf
            '
            'gstrLatXMP
            EXIFInfo = Form1.EXF.ListInfo 'list all tags into the text box
            If EXIFInfo <> "" Then
                If StrComp(EXIFInfo, "File is not in EXIF format.") <> 0 Then
                    If chkGPSLatitude.Value = 1 Then
                        Pos = InStr(1, EXIFInfo, "GPSLatitude:")
                        If Pos <> 0 Then
                            pos1 = InStr(Pos, EXIFInfo, ":")
                            pos2 = InStr(Pos, EXIFInfo, vbCrLf)
                            strGPSLatitude = Mid(EXIFInfo, pos1 + 2, pos2 - pos1)
                            Pos = InStr(1, EXIFInfo, "GPSLatitudeRef: S", vbTextCompare)    'Gerbing 12.10.2016
                            If Pos <> 0 Then
                                strGPSLatitude = "-" & strGPSLatitude
                            End If
                            rstsql.Fields("GPSLatitude") = CDbl(strGPSLatitude)             'Gerbing 12.10.2016
                            rstsql.Update
                        End If
                        If gstrLatXMP <> "" Then                                            'Gerbing 08.04.2019
                            Call Form1.GEOKoordinatenUmrechnenXMP
                            rstsql.Fields("GPSLatitude") = gstrLat
                            rstsql.Update
                        End If
                    End If
                    If chkGPSLongitude.Value = 1 Then
                        Pos = InStr(1, EXIFInfo, "GPSLongitude:")
                        If Pos <> 0 Then
                            pos1 = InStr(Pos, EXIFInfo, ":")
                            pos2 = InStr(Pos, EXIFInfo, vbCrLf)
                            strGPSLongitude = Mid(EXIFInfo, pos1 + 2, pos2 - pos1)
                            Pos = InStr(1, EXIFInfo, "GPSLongitudeRef: W", vbTextCompare)   'Gerbing 12.10.2016
                            If Pos <> 0 Then
                                strGPSLongitude = "-" & strGPSLongitude
                            End If
                            rstsql.Fields("GPSLongitude") = CDbl(strGPSLongitude)           'Gerbing 12.10.2016
                            rstsql.Update
                        End If
                        If gstrLongXMP <> "" Then                                            'Gerbing 08.04.2019
                            Call Form1.GEOKoordinatenUmrechnenXMP
                            rstsql.Fields("GPSlongitude") = gstrLong
                            rstsql.Update
                        End If
                    End If
                    If chkEXIFDateTimeOriginal.Value = 1 Then
                        'Es gelten folgende Prioritäten für die Quelle von EXIFDateTimeOriginal
                        '1.DateTimeOriginal
                        '2.DateTimeDigitized
                        '3.DateTime
                        Pos = InStr(1, EXIFInfo, "DateTimeOriginal")
                        If Pos <> 0 Then
                            pos1 = InStr(Pos, EXIFInfo, vbCrLf)
                            strEXIFDateTimeOriginal = Mid(EXIFInfo, Pos + 18, 19)
                            rstsql.Fields("EXIFDateTimeOriginal") = strEXIFDateTimeOriginal
                            rstsql.Update
                            GoTo rstsqlMoveNext                                                     'Gerbing 11.01.2016
                        End If
                        Pos = InStr(1, EXIFInfo, "DateTimeDigitized")
                        If Pos <> 0 Then
                            pos1 = InStr(Pos, EXIFInfo, vbCrLf)
                            strEXIFDateTimeOriginal = Mid(EXIFInfo, Pos + 19, 19)
                            rstsql.Fields("EXIFDateTimeOriginal") = strEXIFDateTimeOriginal
                            rstsql.Update
                            GoTo rstsqlMoveNext                                                     'Gerbing 11.01.2016
                        End If
                        Pos = InStr(1, EXIFInfo, "DateTime")
                        If Pos <> 0 Then
                            pos1 = InStr(Pos, EXIFInfo, vbCrLf)
                            strEXIFDateTimeOriginal = Mid(EXIFInfo, Pos + 10, 19)
                            rstsql.Fields("EXIFDateTimeOriginal") = strEXIFDateTimeOriginal
                            rstsql.Update
                        End If
                    End If
                End If
            End If
            'Bei mp4 und mov Videos brauche ich die MediaInfo.DLL                                   'Gerbing 18.11.2019
            'If DateinamenErweiterung = "MP4" Or DateinamenErweiterung = "MOV" Then
                If blnMediaInfoInitialized = False Then
                    If Not file_exist(AppPath + "\MediaInfo.dll") Then
                        Msg = LoadResString(1558 + Sprache) & vbCr   'Sorry, the MediaInfo.dll(i386 version) not found in the current path!
                        Msg = Msg & LoadResString(1559 + Sprache) & vbCr 'Put the {MediaInfo.dll} into current path before runnig this Application!
                        Msg = Msg & LoadResString(1560 + Sprache) & vbNewLine & vbNewLine 'We need Mediainfo.dll for checking GPS info in mp4 videos
                        
                        Msg = Msg & LoadResString(1561 + Sprache)   'Do you want to continue without MediaInfo.dll?
                        antwort = MessageBoxW(0, StrPtr(Msg), StrPtr("GERBING Fotosmdb"), vbDefaultButton1 + vbYesNo)
                        If antwort = vbNo Then
                            End                                                                     'Gerbing 18.11.2019
                        Else
                            GoTo rstsqlMoveNext                                                     'Gerbing 05.06.2019
                        End If
                    End If
                    On Error Resume Next
                    Handle = MediaInfo_New()    'hier bekomme ich Laufzeitfehler '48' Datei nicht gefunden mediainfo.dll wenn ich das Programm aus der IDE heraus
                                                'zum zweitenmal starte
                    If Err.Number <> 0 Then
                        MsgBox "invalid DLL-Version(must be i386 version)"
                        End                                                                     'Gerbing 18.11.2019
                    End If
                End If
                Call GetMediaInfo(tempDateiname)
                blnMediaInfoInitialized = True
            'End If
rstsqlMoveNext:
            txtArbeitsfortschritt.Text = tempDateiname                                          'Gerbing 18.11.2019
            rstsql.MoveNext
            DoEvents                                                                            'Gerbing 18.11.2019
        Loop
        rstsql.Close
    Else
        'im schreibgeschützten Zustand nicht möglich
        Msg = gstrFotosMdbLocation & "\Fotos.mdb" & vbNewLine
        'Msg= msg & "Die Datenbank ist schreibgeschützt, Änderungen sind nicht möglich"
        Msg = Msg & LoadResString(2210 + Sprache)
        'MsgBox Msg
        MessageBoxW 0, StrPtr(Msg), StrPtr(LoadResString(1119 + Sprache)), vbInformation
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Call AnpassenNutzerWunsch(Me)                                       'Gerbing 11.03.2017
    Me.Caption = LoadResString(3165 + Sprache)                          'Feld-Aktualisierung durch Import-Wiederholung
    btnStart.tooltipText = LoadResString(3166 + Sprache)                '"Hiermit starten Sie für die aktuelle Datei-Auswahl einen erneuten Metadaten-Import"
    lblArbeitsfortschritt.Caption = LoadResString(1014 + Sprache)       'Arbeitsfortschritt:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If blnMediaInfoInitialized = True Then
        Call MediaInfo_Close(Handle)                                                    'Gerbing 18.11.2019
        Call MediaInfo_Delete(Handle)
    End If
End Sub

Private Sub GetMediaInfo(tempDateiname As String)                                                      'Gerbing 18.11.2019
Dim display As String
    Dim GPS As String
    Dim strLat As String
    Dim strLong As String
    Dim InfoMediadll As String
    Dim strDateTime As String
    Dim Pos As Long
    Dim pos1 As Long
    Dim pos2 As Long
    Dim pos3 As Long
    Dim rc As Long
    Const SW_SHOWNORMAL = 1
    
    rc = MediaInfo_Open(Handle, StrPtr(tempDateiname))
    display = InfoMediadll
    Call MediaInfo_Option(Handle, StrPtr("Complete"), StrPtr(""))
    display = display + StripStrinCtoVB(MediaInfo_Inform(Handle, InformOption_Nothing))
    Call MediaInfo_Close(Handle)
    
    'Die Felder GPSLatitude und GPSLongitude auffüllen
    'xyz Suchen                                                                'Gerbing 18.11.2019
    Pos = InStr(1, display, "xyz")
    If Pos <> 0 Then
        pos1 = InStr(Pos, display, ":")
        pos2 = InStr(pos1, display, "/")
        pos3 = InStr(pos1 + 3, display, "+")
        If pos3 = 0 Then
            pos3 = InStr(pos1 + 3, display, "-")
        End If
        GPS = Mid(display, pos1 + 2, pos2 - pos1 - 2)
        'MsgBox GFfPS
        'zB GPS = "+50.8314+12.8311"
        strLat = Mid(GPS, 1, pos3 - pos1 - 2)
        strLong = Mid(GPS, Len(strLat) + 1, pos2 - pos3)
        rstsql.Fields("GPSLatitude") = strLat
        rstsql.Fields("GPSLongitude") = strLong
    End If
    '---------------------------------------------------------------------------------------------------------------------
    Pos = InStr(1, display, "Encoded Date", vbTextCompare)                                    'das erste "Encoded Date" ist das richtige
    If Pos <> 0 Then
        pos1 = InStr(Pos, display, ": UTC")
        If Pos <> 0 Then
            strDateTime = Mid(display, pos1 + 6, 19)
            strDateTime = Replace(strDateTime, "-", ":")
            rstsql.Fields("ExifDateTimeOriginal") = strDateTime
        End If
    End If
    display = Empty
    Exit Sub
End Sub

