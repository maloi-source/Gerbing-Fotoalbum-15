VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{A8F9B8E7-E699-4FCE-A647-72C877F8E632}#1.8#0"; "editctlsu.ocx"
Begin VB.Form frmVideo 
   BackColor       =   &H00000000&
   ClientHeight    =   9708
   ClientLeft      =   192
   ClientTop       =   216
   ClientWidth     =   12552
   Icon            =   "frmVideo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   9708
   ScaleWidth      =   12552
   StartUpPosition =   1  'Fenstermitte
   WindowState     =   2  'Maximiert
   Begin VB.CommandButton Command1 
      Caption         =   "test video abspielen"
      Height          =   612
      Left            =   9360
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   2172
   End
   Begin EditCtlsLibUCtl.TextBox txtBildbeschreibung 
      Height          =   372
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   7932
      _cx             =   13991
      _cy             =   656
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
      Enabled         =   -1  'True
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
      ReadOnly        =   -1  'True
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
      CueBanner       =   "frmVideo.frx":038A
      Text            =   "frmVideo.frx":03AA
   End
   Begin VB.Label lblLeereForm 
      Caption         =   "Wählen Sie ein anderes Bild, Tasten F2/F3 oder Pfeil-Tasten"
      Height          =   492
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   7812
   End
   Begin WMPLibCtl.WindowsMediaPlayer WMP 
      Height          =   8052
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   11292
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   -1  'True
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   19918
      _cy             =   14203
   End
   Begin VB.Menu mnuDatei 
      Caption         =   "Datei..."
      Visible         =   0   'False
      Begin VB.Menu mnuDateiKopieren 
         Caption         =   "Datei kopieren..."
      End
      Begin VB.Menu mnuGeoPosition 
         Caption         =   "Zeige Geo-Position"
      End
      Begin VB.Menu mnuEinfügenGeoPosition 
         Caption         =   "Einfügen Geo-Position"
      End
      Begin VB.Menu mnuVerknüpfteAnwendung 
         Caption         =   "Öffne die mit der aktuellen Datei  verknüpfte Anwendung"
      End
      Begin VB.Menu mnuÖffneDruckprogramm 
         Caption         =   "Öffne das Druckprogramm für die aktuelle Datei..."
      End
      Begin VB.Menu mnuEmailAnhang 
         Caption         =   "Email mit Anhang senden..."
      End
      Begin VB.Menu mnuExplorer 
         Caption         =   "Öffne ein Explorer-Fenster, wo die aktuelle Datei markiert ist"
      End
      Begin VB.Menu mnuKopiereMdb 
         Caption         =   "Export der aktuell ausgewählten Dateien..."
      End
      Begin VB.Menu mnuImport 
         Caption         =   "Import mit Drag&&Drop..."
      End
      Begin VB.Menu mnuRenammdb 
         Caption         =   "Öffne RenamMdb für die aktuelle Datei"
      End
      Begin VB.Menu mnuWeiterselektieren 
         Caption         =   "Weiterselektieren nur die mit Merkerspalte markierten Dateien anzeigen"
      End
      Begin VB.Menu mnuLöschen 
         Caption         =   "Löschen markierte Dateien(Merkerspalte) in Datenbank und Standort"
      End
      Begin VB.Menu mnuHyperlink 
         Caption         =   "Gehe zum Hyperlink"
      End
      Begin VB.Menu mnuFeldAktualisierung 
         Caption         =   "Feld-Aktualisierung durch Import-Wiederholung..."
      End
      Begin VB.Menu mnuNamenErsetzen 
         Caption         =   "NamenErsetzen AltName && Ort && Situation && Personen"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools..."
      Visible         =   0   'False
      Begin VB.Menu mnuFotosmdbStarten 
         Caption         =   "Fotosmdb starten"
      End
      Begin VB.Menu mnuRenammdbStarten 
         Caption         =   "RenamMdb starten"
      End
   End
   Begin VB.Menu mnuVersion 
      Caption         =   "Version"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuHilfe 
      Caption         =   "Hilfe"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmVideo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim Msg As String
    Dim blnComeFromfrmVideoFormLoad As Boolean

Private Sub Command1_Click()
    WMP.url = AppPath & "\2006\video3.avi"
    WMP.Controls.play
End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim ShiftDown As Boolean
    Dim AltDown As Boolean
    Dim CtrlDown As Boolean
    
    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0

    'Debug.Print "frmVideo.Form_KeyDown Keycode=" & KeyCode & " Shift=" & Shift
    
    If Shift = vbAltMask And KeyCode = 18 Then  '18 = Menu key                      'Gerbing 01.09.2017
        KeyCode = 0
        mnuDatei.Visible = Not mnuDatei.Visible
        mnuTools.Visible = Not mnuTools.Visible
        mnuVersion.Visible = Not mnuVersion.Visible
        mnuHilfe.Visible = Not mnuHilfe.Visible
        Sleep 100  'Sonst flackert es bei Festhalten der Taste Alt                  'Gerbing 05.11.2017
        DoEvents                                                                    'Gerbing 20.01.2018
        Exit Sub
    End If
    
    If KeyCode = vbKeyF1 Or KeyCode = vbKeyF4 Then
        Exit Sub    'F1 und F4 bei Videos wirkungslos                               'Gerbing 26.11.2012
    End If
    '--------------------------------------------------------------------------------------------------
    gblnComefromVideo = True                                                        'Gerbing 16.06.2012
    Call Form1.Form_KeyDown(KeyCode, Shift)
    Select Case KeyCode
        Case vbKeyF6, vbKeyF7, vbKeyF9, vbKeyF10, vbKeyF11
            frmVideo.Show
    End Select
End Sub

Private Sub Form_Load()
    blnComeFromfrmVideoFormLoad = True                                              'Gerbing 04.02.2013
    Call AnpassenNutzerWunsch(Me)                                                   'Gerbing 11.03.2017
    Me.Top = 0                                                                      'Gerbing 20.11.2008
    Me.Left = 0
    'lblLeereForm.Caption = "Wählen Sie ein anderes Video, Tasten F2/F3 oder Alt+Pfeil-Tasten"   'Gerbing 08.11.2005
    lblLeereForm.Caption = LoadResString(1127 + Sprache)
    'mnuxxx.Caption mit LoadResString laden                                         'Gerbing 08.01.2018
    mnuDatei.Caption = LoadResString(3167 + Sprache)
    mnuDateiKopieren.Caption = LoadResString(3168 + Sprache)
    mnuEmailAnhang.Caption = LoadResString(3169 + Sprache)
    mnuExplorer.Caption = LoadResString(3170 + Sprache)
    mnuFeldAktualisierung.Caption = LoadResString(3171 + Sprache)
    mnuFotosmdbStarten.Caption = LoadResString(3172 + Sprache)
    mnuGeoPosition.Caption = LoadResString(3173 + Sprache)
    mnuHilfe.Caption = LoadResString(3174 + Sprache)
    mnuHyperlink.Caption = LoadResString(3175 + Sprache)
    mnuImport.Caption = LoadResString(3176 + Sprache)
    mnuKopiereMdb.Caption = LoadResString(3177 + Sprache)
    mnuLöschen.Caption = LoadResString(3178 + Sprache)
    mnuNamenErsetzen.Caption = LoadResString(3179 + Sprache)
    mnuÖffneDruckprogramm.Caption = LoadResString(3180 + Sprache)
    mnuRenammdb.Caption = LoadResString(3181 + Sprache)
    mnuRenammdbStarten.Caption = LoadResString(3182 + Sprache)
    mnuTools.Caption = LoadResString(3183 + Sprache)
    mnuVerknüpfteAnwendung.Caption = LoadResString(3184 + Sprache)
    mnuVersion.Caption = LoadResString(3185 + Sprache)
    mnuWeiterselektieren.Caption = LoadResString(3186 + Sprache)
    mnuDatei.Visible = False
    mnuTools.Visible = False
    mnuVersion.Visible = False
    mnuHilfe.Visible = False
    
    frmVideo.width = Form1.width                                                    'Gerbing 04.12.2012
    frmVideo.height = Form1.height
    blnComeFromfrmVideoFormLoad = False                                             'Gerbing 04.02.2013
    'Gerbing 02.04.2018
    If Query.chkFensterGrößeÄnderbar.Value = 1 Then                                 'Gerbing 21.05.2012
        'Achtung in der IDE wird unicode in Form.Caption nicht angezeigt
        ShowTitleBar True, False                                     'Gerbing 04.09.2012 'taskbar visible, Video
    Else
        'Achtung in der IDE wird unicode in Form.Caption nicht angezeigt
        ShowTitleBar False, False                                    'Gerbing 04.09.2012 'taskbar unvisible, Video
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        gblnComefromVideo = True                                                        'Gerbing 16.06.2012
        Call Form1.Hilfebox
    End If
End Sub

Public Sub Form_Resize()
    If blnComeFromfrmVideoFormLoad = True Then Exit Sub                                 'Gerbing 04.02.2013
    
    If frmVideo.WindowState = 2 Then    'maximized                                      'Gerbing 04.02.2013
        Form1.WindowState = 2
        Exit Sub
    
    End If
    If frmVideo.WindowState = 1 Then   'minimized                                       'Gerbing 15.01.2013
        Form1.WindowState = 1
        frmGridAndThumb.WindowState = 1                                                 'Gerbing 04.05.2015
        gblnComefromVideo = False
    Else
        frmVideo.WindowState = 0        'normal
        frmVideo.width = glngSaveForm1Width
        frmVideo.height = glngSaveForm1Height
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim retcode As Long                                                             'Gerbing 31.12.2012

    If gblnComeFromBildanzeigen = True Then
        gblnComeFromBildanzeigen = False
        Exit Sub
    End If
    
'    If Form1.lngPointer Then
'        retcode = GdipDisposeImage(Form1.lngPointer)
'        Form1.lngPointer = 0                                                              'Gerbing 19.04.2017
'    End If
'    If m_lngGraphics Then
'        If GdipDeleteGraphics(m_lngGraphics) Then _
'            'MsgBox "Graphics object could not be deleted", vbCritical              'Gerbing 30.11.2016
'        End If
'    End If
'    GdiplusShutdown m_lngInstance
    
    'Unload Hilfebx
    Set Form1.EXF = Nothing                       'Gerbing 07.05.2007
    Unload frmGridAndThumb
    Unload Hilfebx
    Unload KommentarForm
    Unload Query
    'Unload QueryJedesFeld
    Unload MP
    Unload Form1
    End
End Sub


Private Sub mnuDateiKopieren_Click()
    frmZwischenablageOderOrdner.Show 1                                                      'Gerbing 11.08.2017
End Sub

Private Sub mnuFotosmdbStarten_Click()
    Dim AppId
    Dim Msg As String
    Dim cmdline As String
    
    'If Dir(AppPath & "\fotosmdb.exe") = "" Then
    If file_path_exist(AppPath & "\fotosmdb.exe") = False Then
        'msg = "Fotosmdb konnte nicht gestartet werden." & vbNewLine
        Msg = LoadResString(2167 + Sprache) & vbNewLine
        'msg = msg & "Fotosmdb.exe muss im gleichen Ordner stehen wie fotos.exe"
        Msg = Msg & LoadResString(2168 + Sprache)
        'MsgBox Msg
        MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
        Exit Sub
    End If
    'CommandLine aufbauen mit access
        'fotosmdblocation=...;
    
    'CommandLine aufbauen mit sql server
        'sqlservername=...;
        'datenbankname=...;
        'WindowsAuthentication=0; heißt nein
        'WindowsAuthentication=1; heißt ja
        'username=...;
        'Password=...;
        'StandortFotos=...;

    'CommandLine aufbauen mit access
    If gblnSQLServerVersion = False Then
        If gstrFotosMdbLocation <> "" Then                                                          'Gerbing 07.11.2011
            AppId = Shell(AppPath & "\fotosmdb.exe" & " " & "fotosmdblocation=" & gstrFotosMdbLocation & ";", vbNormalFocus)
            AppActivate AppId
        Else
            AppId = Shell(AppPath & "\fotosmdb.exe", vbNormalFocus)
            AppActivate AppId
        End If
    Else
    'CommandLine aufbauen mit sql server
        cmdline = "sqlservername=" & PublicSQLServer & ";"
        cmdline = cmdline & "datenbankname=" & PublicSQLDatabase & ";"
        cmdline = cmdline & "WindowsAuthentication=" & PublicWindowsAuthentication & ";"
        If PublicWindowsAuthentication = "0" Then
            cmdline = cmdline & "username=" & PublicSQLServerUserName & ";"
            cmdline = cmdline & "Password=" & PublicSQLServerPassword & ";"
        End If
        cmdline = cmdline & "StandortFotos=" & PublicLocationFotos & ";"
        AppId = Shell(AppPath & "\fotosmdb.exe" & " " & cmdline, vbNormalFocus)
        AppActivate AppId
    End If

End Sub

Private Sub mnuHilfe_Click()                                                    'Gerbing 01.09.2017
    Dim RetVal As Long
    Dim CHMFile As String

    If Sprache = 0 Then                             'Gerbing 08.11.2005
        CHMFile = AppPath & "\Help\Deutsch\fotos.CHM"                           'Gerbing 14.03.2007
        If isUnicodeString(CHMFile) = True Then
            'Msg = "CHM-Help-Dateien lassen sich im Unicode-Pfad nicht öffnen, das hat Microsoft nicht vorgesehen" & vbNewLine
            'Msg = Msg & "Kopieren Sie die CHM-Help-Dateien in einen Pfad ohne Unicode-Zeichen"
            Msg = CHMFile & vbNewLine
            Msg = Msg & LoadResString(2544 + Sprache) & vbNewLine
            Msg = Msg & LoadResString(2545 + Sprache)
            MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
            Exit Sub
        Else
            RetVal = RunShellExecute(Me.hWnd, "open", CHMFile, vbNull, vbNull, 1)
            If RetVal <= 32 Then
                Call HelpFileErrorMsg(RetVal, CHMFile)
            End If
        End If
    Else
        CHMFile = AppPath & "\Help\English\fotos.CHM"                           'Gerbing 14.03.2007
        If isUnicodeString(CHMFile) = True Then
            'Msg = "CHM-Help-Dateien lassen sich im Unicode-Pfad nicht öffnen, das hat Microsoft nicht vorgesehen" & vbNewLine
            'Msg = Msg & "Kopieren Sie die CHM-Help-Dateien in einen Pfad ohne Unicode-Zeichen"
            Msg = CHMFile & vbNewLine
            Msg = Msg & LoadResString(2544 + Sprache) & vbNewLine
            Msg = Msg & LoadResString(2545 + Sprache)
            MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
            Exit Sub
        Else
            RetVal = RunShellExecute(Me.hWnd, "open", CHMFile, vbNull, vbNull, 1)
            If RetVal <= 32 Then
                Call HelpFileErrorMsg(RetVal, CHMFile)
            End If
        End If
    End If
End Sub

Private Sub mnuImport_Click()
    If gblnSQLServerVersion = True Then
        MsgBox "Diese Funktion wird beim SQL-Server nicht unterstützt"
        MsgBox LoadResString(1829 + Sprache)
    Else
        Call Form1.MediaPlayerStop
        ImportForm.Show 1
        frmVideo.lblLeereForm.Visible = True
    End If
End Sub

Private Sub mnuKopiereMdb_Click()
    Call Form1.MediaPlayerStop
    'me.MousePointer = vbDefault                                                'Gerbing 29.07.2007
    Form1.blnUnloadExportForm = False
    ExportForm.Show 1
    frmVideo.lblLeereForm.Visible = True                                            'Gerbing 04.12.2012
End Sub

Private Sub WMP_DeviceSyncError(ByVal pDevice As WMPLibCtl.IWMPSyncDevice, ByVal pMedia As Object)
    MsgBox "WMP_DeviceSyncError"
End Sub

Private Sub WMP_Error()
    If WMP.url <> "" Then                                                          'Gerbing 29.08.2008
        Msg = WMP.url & NL                                                         'Gerbing 29.08.2008
'        Msg = Msg & "Es ist ein Fehler beim Abspielen der Datei aufgetreten." & NL
'        Msg = Msg & "Kontrollieren Sie ob die Pfadangabe richtig ist." & NL
'        Msg = Msg & "Kontrollieren Sie, ob sich die Datei außerhalb von diesem Programm abspielen läßt." & NL & NL
        Msg = Msg & LoadResString(2283 + Sprache) & NL                              'Gerbing 09.03.2014
        Msg = Msg & LoadResString(2284 + Sprache) & NL
        Msg = Msg & LoadResString(2285 + Sprache) & NL & NL
        'MsgBox Msg
        MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
        lblLeereForm.Visible = True
        Form1.Hide                                                                  'Gerbing 15.02.2014
    End If
    Me.MousePointer = vbDefault
End Sub

Private Sub WMP_MediaError(ByVal pMediaObject As Object)
    MakeGradient Me, vbBlue, vbGreen, GRADIENT_FILL_RECT_V                      'Gerbing 15.02.2014
    MsgBox "MediaPlayer1_MediaError"
    If Form1.F6Continous Then
        lblLeereForm.Visible = True
        Form1.blnTimer1Enabled = True
        EnableTimer Form1.lngTimer1Interval
    End If
End Sub

Private Sub WMP_MouseDown(ByVal nButton As Integer, ByVal nShiftState As Integer, ByVal fX As Long, ByVal fY As Long)
    If nButton = vbRightButton Then                                              'Gerbing 26.11.2012
        Call Form1.Hilfebox
    End If
End Sub

Private Sub WMP_PlayStateChange(ByVal NewState As Long)                         'Gerbing 01.05.2013
    Dim blnMustUpdate As Boolean

    'player .playState
    'Possible Values
    'This property is a read-only Number (long).
    'Value   State   Description
    '0   Undefined   Windows Media Player is in an undefined state.
    '1   Stopped     Playback of the current media item is stopped.
    '2   Paused      Playback of the current media item is paused. When a media item is paused, resuming playback begins from the same location.
    '3   Playing     The current media item is playing.
    '4   ScanForward The current media item is fast forwarding.
    '5   ScanReverse The current media item is fast rewinding.
    '6   Buffering       The current media item is getting additional data from the server.
    '7   Waiting     Connection is established, but the server is not sending data. Waiting for session to begin.
    '8   MediaEnded  Media item has completed playback.
    '9   Transitioning   Preparing new media item.
    '10  Ready       Ready to begin playing.
    '11  Reconnecting    Reconnecting to stream.

    'Debug.Print "NewState=" & NewState
    If gblnComeFromBeenden = True Then Exit Sub                                 'Gerbing 30.11.2018
    blnMustUpdate = False
    If NewState = 3 Then                                                        '3=playing
        glngStartMillisek = timeGetTime                                         'Gerbing 30.05.2019
        If IsNull(frmGridAndThumb.Adodc1.Recordset(LoadResString(1106 + Sprache))) Then  '30.11.2018 hier kam run time error
            frmGridAndThumb.Adodc1.Recordset(LoadResString(1106 + Sprache)) = WMP.currentMedia.imageSourceWidth
            blnMustUpdate = True
        End If
        If IsNull(frmGridAndThumb.Adodc1.Recordset(LoadResString(1107 + Sprache))) Then
            frmGridAndThumb.Adodc1.Recordset(LoadResString(1107 + Sprache)) = WMP.currentMedia.imageSourceHeight
            blnMustUpdate = True
        End If
        On Error Resume Next
        If IsNull(frmGridAndThumb.Adodc1.Recordset("VideoDuration")) Then
            frmGridAndThumb.Adodc1.Recordset("VideoDuration") = WMP.currentMedia.duration
            blnMustUpdate = True
        End If
        Err.Number = 0
        If blnMustUpdate = True Then
            frmGridAndThumb.Adodc1.Recordset.Update
            If Err.Number <> 0 Then                                             'Gerbing 30.05.2012
                Msg = Err.Number & vbNewLine
                Msg = Msg & Err.Description
                'MsgBox Msg
                MessageBoxW 0, StrPtr(Msg), StrPtr(LoadResString(1119 + Sprache)), vbInformation
            End If
        End If
        On Error GoTo 0
    End If
    If NewState = 8 Then                                                        '8=MediaEnded 'Gerbing 07.05.2013
        glngEndMillisek = timeGetTime                                           'Gerbing 30.05.2019
        If glngEndMillisek - glngStartMillisek < 300 Then                       'Gerbing 30.05.2019
            Debug.Print "Mediaplayer Ended nach millisekunden=" & glngEndMillisek - glngStartMillisek 'Gerbing 30.05.2019
            Call WMP_Error                                                      'Gerbing 30.05.2019
        End If                                                                  'Gerbing 30.05.2019
        If Form1.F6Continous = True Then
            Call Form1.GeheEinBildVorwärts
            On Error Resume Next
            Call Form1.BildAnzeigen
            On Error GoTo 0
        End If
    End If
    If NewState = 1 Then                                                        '1=stopped 'Gerbing 24.10.2013
        WMP.Controls.play
    End If
End Sub

Private Sub WMP_Warning(ByVal WarningType As Long, ByVal Param As Long, ByVal Description As String)
    MsgBox Description
End Sub

Private Sub mnuEmailAnhang_Click()
    Call Form1.EmailMitAnhangSenden                                                               'Gerbing 01.09.2017
End Sub

Private Sub mnuExplorer_Click()                                                             'Gerbing 01.09.2017
    Dim RetVal As Long
    
    If gblnWasOptThumbClick = False Then                                                    'Gerbing 24.11.2016
        gstrFRODN = Replace(gstrRowColChangeName, "+:\", gstrFotosMdbLocation & "\")        'Gerbing 04.01.2006
    End If
    If gstrFRODN = "" Then                                                              'Gerbing 09.07.2008
        'wenn es noch kein Ereignis DbGridNeu_RowColChange gab, dann ist gstrRowColChangeName und damit gstrFRODN leer
        gstrFRODN = frmGridAndThumb.Adodc1.Recordset.Fields(LoadResString(1028 + Sprache))
        gstrFRODN = Replace(gstrFRODN, "+:\", gstrFotosMdbLocation & "\")
    End If

    'Öffne ein Explorer-Fenster, wo die aktuelle Datei markiert ist                         Gerbing 12.11..2007
    'Hierbei gibt es einen Fehler wenn im Dateiname ein Komma
    'enthalten ist -> "Der Pfad '...Teil hinter dem Komma' ist nicht vorhanden oder weist auf kein
    'Verzeichnis
    'Man muss den Dateinamen in doppelte Hochkomma einschließen
    'RetVal = RunShellExecute(Me.hWnd, "open", "explorer.exe", "/e,/select," & """" & gstrFRODN & """", vbNull, 1) 'Gerbing 12.11.2007
    RetVal = RunShellExecute(Me.hWnd, "open", "explorer.exe", "/e,/select," & """" & gstrFRODN & """", vbNullString, 1) 'Gerbing 31.12.2007
    'IPTCPresent = False setzen
    frmGridAndThumb.Adodc1.Recordset("IPTCPresent") = 0                                     'Gerbing 13.12.2016
End Sub

Private Sub mnuFeldAktualisierung_Click()                                                       'Gerbing 01.09.2017
    Dim SQL As String
    Dim rc As Integer
    
    'das geht nur für JPG files, in anderen ist keine GEO-Positionen vorhanden
    'und nur wenn es die Felder GPSLatitude und GPSLongitude gibt
    If Not (gblnVollversion = True And gblnProversion = True) Then                          'Gerbing 27.09.2016
        Msg = LoadResString(2335 + Sprache) 'Für diese Funktion benötigen Sie die Professional Version.
        MsgBox Msg
        Exit Sub
    End If
    'Kontrollieren, ob es die Felder GPSLatitude und GPSLongitude in der Tabelle fotos gibt 'Gerbing 05.09.2016
    'wenn nicht, MsgBox zeigen und Abbrechen
    
    rc = Form1.GPSFelderPrüfen                                                                  'Gerbing 02.10.2019
    If rc = 0 Then Exit Sub                                                                     'Gerbing 02.10.2019
    frmFeldAktualisierung.Show 1
End Sub

Private Sub mnuGeoPosition_Click()                                                              'Gerbing 01.09.2017
    Call Form1.ZeigeGEOPosition                                                                       'Gerbing 03.10.2016
End Sub

Private Sub mnuEinfügenGeoPosition_Click()
    'das Einfügen der Geo-Position geht nicht nur für JPG files sondern für alle
    'aber nur wenn es die Felder GPSLatitude und GPSLongitude gibt
    Dim rc As Integer
    
    rc = Form1.GPSFelderPrüfen                                                                  'Gerbing 02.10.2019
    If rc = 0 Then Exit Sub                                                                     'Gerbing 02.10.2019
    frmGPSInDatenbankEintragen.Show 1                                                           'Gerbing 02.10.2019
End Sub

Private Sub mnuHyperlink_Click()                                                                'Gerbing 01.09.2017
    Dim strTemp As String
    Dim RetVal As Long
    
    If Not (gblnVollversion = True And gblnProversion = True) Then                              'Gerbing 27.09.2016
        Msg = LoadResString(2335 + Sprache) 'Für diese Funktion benötigen Sie die Professional Version.
        MsgBox Msg
        Exit Sub
    End If
    If Right(frmGridAndThumb.DBGridNeu.Text, 1) = "#" And Left(frmGridAndThumb.DBGridNeu.Text, 1) = "#" Then                    '# muss da sein wenn der Hyperlink auf die eiegen Festplatte verweist
        If Mid(frmGridAndThumb.DBGridNeu.Text, 2, 3) = "+:\" Then
            strTemp = Replace(frmGridAndThumb.DBGridNeu.Text, "+:\", gstrFotosMdbLocation & "\")                'Gerbing 07.11.2011
            strTemp = Mid(strTemp, 2, Len(strTemp) - 2)
        Else
            strTemp = Mid(frmGridAndThumb.DBGridNeu.Text, 2, Len(frmGridAndThumb.DBGridNeu.Text) - 2)
        End If
        RetVal = RunShellExecute(Me.hWnd, "Open", strTemp, "", gstrFotosMdbLocation, 1)
        If RetVal <= 32 Then
            MsgBox LoadResString(3129 + Sprache) 'Es wurde kein geeigneter Browser gefunden um diese URL zu öffnen
        End If
    Else
        If frmGridAndThumb.DBGridNeu.Text <> "" Then                                                            'Gerbing 27.09.2016
            RetVal = RunShellExecute(Me.hWnd, "Open", frmGridAndThumb.DBGridNeu.Text, "", gstrFotosMdbLocation, 1)
            If RetVal <= 32 Then
                'MsgBox LoadResString(3129 + Sprache) 'Es wurde kein geeigneter Browser gefunden um diese URL zu öffnen
'                        Msg = "Die gewählte Aktion erfordert, dass das aktive Feld einen Hyperlink enthält," & vbNewLine
'                        Msg = Msg & "im Format #Hyperlink#"
'                        Msg = Msg & "Inhalt des aktiven Feldes=" & frmGridAndThumb.DBGridNeu.Text
                Msg = LoadResString(3129 + Sprache) & vbNewLine
                Msg = Msg & LoadResString(3126 + Sprache) & vbNewLine
                'Msg = Msg & LoadResString(3127 + Sprache) & vbNewLine                              'Gerbing 27.09.2016 auskommentiert
                Msg = Msg & LoadResString(3128 + Sprache) & frmGridAndThumb.DBGridNeu.Text
                'MsgBox Msg
                MessageBoxW 0, StrPtr(Msg), StrPtr(LoadResString(1119 + Sprache)), vbInformation
            End If
            Exit Sub
        End If
'            Msg = "Die gewählte Aktion erfordert, dass das aktive Feld einen Hyperlink enthält," & vbNewLine
'            Msg = Msg & "im Format #Hyperlink#"
'            Msg = Msg & "Inhalt des aktiven Feldes=" & frmGridAndThumb.DBGridNeu.Text
        Msg = LoadResString(3126 + Sprache) & vbNewLine
        'Msg = Msg & LoadResString(3127 + Sprache) & vbNewLine                              'Gerbing 27.09.2016
        'Msg = Msg & LoadResString(3128 + Sprache) & frmGridAndThumb.DBGridNeu.Text
        'MsgBox Msg
        MessageBoxW 0, StrPtr(Msg), StrPtr(LoadResString(1119 + Sprache)), vbInformation
        Exit Sub
    End If
End Sub

Private Sub mnuLöschen_Click()                                                              'Gerbing 01.09.2017
    Dim antwort As Long
    Dim KeyCode As Integer
    Dim Shift As Integer
    Dim SQL As String
    
    SQL = " SELECT *"
'    SQL = SQL & " FROM " & "Fotos"
'    SQL = SQL & " WHERE Merker<>0
'    SQL = SQL & " ORDER BY Dateiname" & ";"
    SQL = " SELECT *"
    SQL = SQL & " FROM Fotos"
    SQL = SQL & " WHERE " & LoadResString(2524 + Sprache) & "<>0"
    SQL = SQL & " ORDER BY " & LoadResString(1028 + Sprache) & ";"
    On Error Resume Next                                                            'Gerbing 01.09.2017
    rstsql.Close
    On Error GoTo 0
    With rstsql
        .Source = SQL
        .ActiveConnection = DBado                                                       'Gerbing 23.11.2017
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    If rstsql.EOF Then
        rstsql.Close
        MsgBox LoadResString(3061 + Sprache) 'Es gibt keine mit Merkerspalte markierten Sätze
        Exit Sub
    End If
    'msg = "Anzahl markierte Dateien = " & rst1.RecordCount                         'Gerbing 25.06.2008
    'msg = msg & "Wollen Sie wirklich alle mit der Merkerspalte markierten Dateien aus der Datenbank und an ihrem Standort löschen?" & vbnewline
    'msg = msg & "Sie gelangen anschließend in das Fenster zum Angeben der Suchkriterien."
    Msg = LoadResString(2274 + Sprache) & rstsql.RecordCount & vbNewLine              'Gerbing 25.06.2008
    Msg = Msg & LoadResString(1523 + Sprache) & vbNewLine
    Msg = Msg & LoadResString(1524 + Sprache)
    'antwort = MsgBox(msg, vbDefaultButton1 + vbYesNo)
    antwort = MsgBox(Msg, vbDefaultButton1 + vbYesNo)
    If antwort = vbNo Then
        Exit Sub
    End If
    Do Until rstsql.EOF
        Call Form1.LöschenInDatenbankUndStandort(rstsql.Fields(LoadResString(1028 + Sprache)), rstsql) '1028=Dateiname
        rstsql.MoveNext
    Loop
    rstsql.Close
    'so tun als wäre F8 gedrückt worden
    KeyCode = vbKeyF8
    Shift = 0
    'Tastatur-Eingabe weiterreichen
    Sleep (3000)
    Call Form1.Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub mnuNamenErsetzen_Click()                                                        'Gerbing 01.09.2017
    Dim antwort As Long
    Dim strPath As String                                                                   'Gerbing 11.04.2017
    Dim strFile As String                                                                   'Gerbing 11.04.2017
    Dim strOrt As String                                                                    'Gerbing 11.04.2017
    Dim strSituation As String                                                              'Gerbing 11.04.2017
    Dim strPersonen As String                                                               'Gerbing 11.04.2017
    Dim DateinameNeuMitPlus As String                                                       'Gerbing 11.04.2017
    Dim rc As Long                                                                          'Gerbing 11.04.2017
    Dim ThumbnameAlt As String                                                              'Gerbing 11.04.2017
    Dim ThumbnameNeu As String                                                              'Gerbing 11.04.2017
    Dim DateinameAlt As String                                                              'Gerbing 11.04.2017
    Dim DateinameNeu As String                                                              'Gerbing 11.04.2017
    Dim DateinamenErweiterung As String
    Dim pos As Long
    Dim SQL As String
    
    'Es gibt viele nichtssagende Dateinamen wie Juni028.jpg Chemnitz23.jpg Ostern014.jpg
    'Für alle Fotos im Suchergebnis soll der Dateiname zusammengesetzt werden aus
    'NameAlt & Ort(wenn vorhanden) & Situation(wenn vorhanden) & Personen(wenn vorhanden)
    'genauso soll das Feld Dateiname und DateinameKurz geändert werden
    'ebenso der Name in GerbingThumbs
    Screen.MousePointer = vbHourglass
    If gblnSchreibgeschützt = False Then
        SQL = frmGridAndThumb.rsDataGrid.Source
        On Error Resume Next                                                                'Gerbing 24.08.2017
        rstsql.Close
        Err.Number = 0
        With rstsql
            .Source = SQL
            .ActiveConnection = DBado                                                       'Gerbing 23.11.2017
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
        Msg = "Anzahl zu ändernde Dateinamen=" & rstsql.RecordCount & vbNewLine
        Msg = Msg & "Wollen Sie diese wirklich ändern?" & vbNewLine
        antwort = MsgBox(Msg, vbYesNo + vbDefaultButton1)
        If antwort = vbNo Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        On Error GoTo 0                                                                     'Gerbing 24.08.2017
        Do Until rstsql.EOF
            DateinameNeu = ""
            DateinameAlt = Replace(rstsql.Fields(LoadResString(1028 + Sprache)), "+:\", gstrFotosMdbLocation & "\")
            'file_split splits a complete file name into directory, file name and extension:
            Call file_split(DateinameAlt, strPath, strFile, DateinamenErweiterung)
            If Not IsNull(rstsql.Fields(LoadResString(1025 + Sprache))) And rstsql.Fields(LoadResString(1025 + Sprache)) <> "" Then        '1025=Ort
                strOrt = rstsql.Fields(LoadResString(1025 + Sprache))
                DateinameNeu = DateinameNeu & strOrt
            End If
            If Not IsNull(rstsql.Fields(LoadResString(1024 + Sprache))) And rstsql.Fields(LoadResString(1024 + Sprache)) <> "" Then        '1024=Situation
                strSituation = rstsql.Fields(LoadResString(1024 + Sprache))
                DateinameNeu = DateinameNeu & strSituation
            End If
            If Not IsNull(rstsql.Fields(LoadResString(1027 + Sprache))) And rstsql.Fields(LoadResString(1027 + Sprache)) <> "" Then        '1027=Personen
                strPersonen = rstsql.Fields(LoadResString(1027 + Sprache))
                DateinameNeu = DateinameNeu & strPersonen
            End If
            'Prüfen, ob verbotene Zeichen vorkommen diese werden ersetzt durch "-"
            'Ein Dateiname darf keines der folgenden Zeichen enthalten: \/:*?""<>|
            pos = InStr(DateinameNeu, "\")
            If pos <> 0 Then
                DateinameNeu = Replace(DateinameNeu, "\", "-")
            End If
            pos = InStr(DateinameNeu, "/")
            If pos <> 0 Then
                DateinameNeu = Replace(DateinameNeu, "/", "-")
            End If
            pos = InStr(DateinameNeu, ":")
            If pos <> 0 Then
                DateinameNeu = Replace(DateinameNeu, ":", "-")
            End If
            pos = InStr(DateinameNeu, "*")
            If pos <> 0 Then
                DateinameNeu = Replace(DateinameNeu, "*", "-")
            End If
            pos = InStr(DateinameNeu, "?")
            If pos <> 0 Then
                DateinameNeu = Replace(DateinameNeu, "?", "-")
            End If
            pos = InStr(DateinameNeu, """")
            If pos <> 0 Then
                DateinameNeu = Replace(DateinameNeu, """", "-")
            End If
            pos = InStr(DateinameNeu, "<")
            If pos <> 0 Then
                DateinameNeu = Replace(DateinameNeu, "<", "-")
            End If
            pos = InStr(DateinameNeu, ">")
            If pos <> 0 Then
                DateinameNeu = Replace(DateinameNeu, ">", "-")
            End If
            pos = InStr(DateinameNeu, "|")
            If pos <> 0 Then
                DateinameNeu = Replace(DateinameNeu, "|", "-")
            End If
            If DateinameNeu = "" Then GoTo rstsqlMoveNext
            DateinameNeu = strFile & "_" & DateinameNeu
            Debug.Print DateinameNeu
            
            rstsql.Fields(LoadResString(1031 + Sprache)) = DateinameNeu & "." & DateinamenErweiterung   '1031=DateinameKurz
            DateinameNeuMitPlus = Replace(strPath & DateinameNeu & "." & DateinamenErweiterung, gstrFotosMdbLocation & "\", "+:\")
            rstsql.Fields(LoadResString(1028 + Sprache)) = DateinameNeuMitPlus                          '1028=Dateiname
            rc = NameAs(DateinameAlt, strPath & DateinameNeu & "." & DateinamenErweiterung)
            'jetzt untersuchen, ob rename altername As neuername erfolgreich war
            If rc = 72 Then 'Kann die datei nicht erzeugen, wenn es diese bereits gibt
                Msg = "Rename war nicht erfolgreich" & vbNewLine
                Msg = Msg & "altname=" & DateinameAlt & vbNewLine
                Msg = Msg & "neuname=" & strPath & DateinameNeu & "." & DateinamenErweiterung & vbNewLine
                Msg = Msg & "rc=" & rc & vbNewLine
                Msg = Msg & "Machen Sie die Änderung mit RenamMdb.exe" & vbNewLine
                Msg = Msg & "das Programm wird beendet"
                MsgBox Msg
                End
            End If
            ThumbnameAlt = strPath & "GerbingThumbs\" & strFile & "." & DateinamenErweiterung & ".jpg"
            ThumbnameNeu = strPath & "GerbingThumbs\" & DateinameNeu & "." & DateinamenErweiterung & ".jpg"
            rc = NameAs(ThumbnameAlt, ThumbnameNeu)
            rstsql.Update
rstsqlMoveNext:
            rstsql.MoveNext
        Loop
        rstsql.Close
    Else
        'im schreibgeschützten Zustand nicht möglich
        Msg = gstrFotosMdbLocation & "\Fotos.mdb" & vbNewLine
        'Msg= msg & "Die Datenbank ist schreibgeschützt, Änderungen sind nicht möglich"
        Msg = Msg & LoadResString(2210 + Sprache)       '2210=Die Datenbank ist schreibgeschützt, Änderungen sind nicht möglich
        'MsgBox Msg
        MessageBoxW 0, StrPtr(Msg), StrPtr(LoadResString(1119 + Sprache)), vbInformation    '1119=GERBING Fotoalbum 15
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    MsgBox "Fertig - Programm wird beendet"
    End
End Sub

Private Sub mnuÖffneDruckprogramm_Click()
    If gblnWasOptThumbClick = False Then                                                    'Gerbing 24.11.2016
        gstrFRODN = Replace(gstrRowColChangeName, "+:\", gstrFotosMdbLocation & "\")        'Gerbing 04.01.2006
    End If
    If gstrFRODN = "" Then                                                              'Gerbing 09.07.2008
        'wenn es noch kein Ereignis DbGridNeu_RowColChange gab, dann ist gstrRowColChangeName und damit gstrFRODN leer
        gstrFRODN = frmGridAndThumb.Adodc1.Recordset.Fields(LoadResString(1028 + Sprache))
        gstrFRODN = Replace(gstrFRODN, "+:\", gstrFotosMdbLocation & "\")
    End If

    Shell ("rundll32.exe SHELL32,OpenAs_RunDLL " & gstrFRODN)                               'Gerbing 07.08.2013 01.09.2017
End Sub

Private Sub mnuRenamMdb_Click()                                                             'Gerbing 01.09.2017
    '3181 = Öffne RenamMdb für dier aktuelle Datei
    Dim AppId
    Dim cmdline As String                                                                   'Gerbing 07.11.2011

    If file_path_exist(AppPath & "\RenamMdb.exe") = False Then
        'msg = "RenamMdb konnte nicht gestartet werden." & vbNewLine
        Msg = LoadResString(2169 + Sprache) & vbNewLine
        'msg = msg & "RenamMdb.exe muss im gleichen Ordner stehen wie fotos.exe"
        Msg = Msg & LoadResString(2170 + Sprache)
        MsgBox Msg
        Exit Sub
    End If
    
    'CommandLine aufbauen mit access
        'RowColChangeName=...;                                                  'Gerbing 09.10.2014
        'fotosmdblocation=...;
    
    'CommandLine aufbauen mit sql server
        'RowColChangeName=...;                                                  'Gerbing 09.10.2014
        'sqlservername=...;
        'datenbankname=...;
        'WindowsAuthentication=0; heißt nein
        'WindowsAuthentication=1; heißt ja
        'username=...;
        'Password=...;
        'StandortFotos=...;

    'CommandLine aufbauen mit access
    If gblnSQLServerVersion = False Then
        If gstrRowColChangeName = "" Then                                       'Gerbing 09.10.2014
            gstrRowColChangeName = frmGridAndThumb.Adodc1.Recordset.Fields(LoadResString(1028 + Sprache))
        End If
        cmdline = "RowColChangeName=" & gstrRowColChangeName & ";"
        If gstrFotosMdbLocation <> "" Then
            cmdline = cmdline & "fotosmdblocation=" & gstrFotosMdbLocation & ";"
        End If
        AppId = Shell(AppPath & "\RenamMdb.exe " & cmdline, vbNormalFocus)
        AppActivate AppId
    Else
    'CommandLine aufbauen mit sql server
        cmdline = "sqlservername=" & PublicSQLServer & ";"
        cmdline = cmdline & "datenbankname=" & PublicSQLDatabase & ";"
        cmdline = cmdline & "WindowsAuthentication=" & PublicWindowsAuthentication & ";"
        If PublicWindowsAuthentication = "0" Then
            cmdline = cmdline & "username=" & PublicSQLServerUserName & ";"
            cmdline = cmdline & "Password=" & PublicSQLServerPassword & ";"
        End If
        cmdline = cmdline & "StandortFotos=" & PublicLocationFotos & ";"
        cmdline = cmdline & "gstrRowColChangeName=" & gstrRowColChangeName & ";"
        AppId = Shell(AppPath & "\RenamMdb.exe" & " " & cmdline, vbNormalFocus)
        AppActivate AppId
    End If
End Sub

Private Sub mnuRenammdbStarten_Click()
    Dim AppId
    Dim Msg As String
    Dim cmdline As String
    
    'If Dir(AppPath & "\RenamMdb.exe") = "" Then
    If file_path_exist(AppPath & "\RenamMdb.exe") = False Then
        'msg = "RenamMdb konnte nicht gestartet werden." & vbNewLine
        Msg = LoadResString(2169 + Sprache) & vbNewLine
        'msg = msg & "RenamMdb.exe muss im gleichen Ordner stehen wie fotos.exe"
        Msg = Msg & LoadResString(2170 + Sprache)
        'MsgBox Msg
        MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
        Exit Sub
    End If
    'CommandLine aufbauen mit access
        'fotosmdblocation=...;
    
    'CommandLine aufbauen mit sql server
        'sqlservername=...;
        'datenbankname=...;
        'WindowsAuthentication=0; heißt nein
        'WindowsAuthentication=1; heißt ja
        'username=...;
        'Password=...;
        'StandortFotos=...;

    'CommandLine aufbauen mit access
    If gblnSQLServerVersion = False Then
        If gstrFotosMdbLocation <> "" Then                                                          'Gerbing 07.11.2011
            AppId = Shell(AppPath & "\RenamMdb.exe" & " " & "fotosmdblocation=" & gstrFotosMdbLocation & ";", vbNormalFocus)
            AppActivate AppId
        Else
            AppId = Shell(AppPath & "\RenamMdb.exe", vbNormalFocus)
            AppActivate AppId
        End If
    Else
    'CommandLine aufbauen mit sql server
        cmdline = "sqlservername=" & PublicSQLServer & ";"
        cmdline = cmdline & "datenbankname=" & PublicSQLDatabase & ";"
        cmdline = cmdline & "WindowsAuthentication=" & PublicWindowsAuthentication & ";"
        If PublicWindowsAuthentication = "0" Then
            cmdline = cmdline & "username=" & PublicSQLServerUserName & ";"
            cmdline = cmdline & "Password=" & PublicSQLServerPassword & ";"
        End If
        cmdline = cmdline & "StandortFotos=" & PublicLocationFotos & ";"
        AppId = Shell(AppPath & "\RenamMdb.exe" & " " & cmdline, vbNormalFocus)
        AppActivate AppId
    End If
End Sub

Private Sub mnuVerknüpfteAnwendung_Click()                                                  'Gerbing 01.09.2017
    Dim RetVal As Long
    Dim intLänge As Integer
    Dim ErrorText As String
    Dim DateinamenErweiterung As String
    
    If gblnWasOptThumbClick = False Then                                                    'Gerbing 24.11.2016
        gstrFRODN = Replace(gstrRowColChangeName, "+:\", gstrFotosMdbLocation & "\")        'Gerbing 04.01.2006
    End If
    If gstrFRODN = "" Then                                                              'Gerbing 09.07.2008
        'wenn es noch kein Ereignis DbGridNeu_RowColChange gab, dann ist gstrRowColChangeName und damit gstrFRODN leer
        gstrFRODN = frmGridAndThumb.Adodc1.Recordset.Fields(LoadResString(1028 + Sprache))
        gstrFRODN = Replace(gstrFRODN, "+:\", gstrFotosMdbLocation & "\")
    End If
    '--------------------------------------------------------------------
    RetVal = RunShellExecute(Me.hWnd, vbNullString, gstrFRODN, vbNullString, vbNullString, 1)    'Gerbing 18.01.2014
    If RetVal <= 32 Then
        If gstrRowColChangeName = "" Then                                                   'Gerbing 10.11.2016
            gstrRowColChangeName = gstrFRODN
        End If
        If Mid(gstrRowColChangeName, Len(gstrRowColChangeName) - 3, 1) = "." Then           'Gerbing 25.06.2006
            intLänge = 3
        End If
        If Mid(gstrRowColChangeName, Len(gstrRowColChangeName) - 4, 1) = "." Then
            intLänge = 4
        End If
        If Mid(gstrRowColChangeName, Len(gstrRowColChangeName) - 5, 1) = "." Then
            intLänge = 5
        End If
        DateinamenErweiterung = Right(gstrRowColChangeName, intLänge)
        ErrorText = GetShellError(RetVal)           'Gerbing 20.08.2008
        Msg = "Errortext=" & ErrorText & vbNewLine
        Msg = Msg & "Errornr=" & RetVal & vbNewLine & vbNewLine
        
        'Msg = "Der Dateiname lautet nach Ersetzen von +:\ folgendermaßen:" & vbNewLine
        Msg = Msg & LoadResString(1375 + Sprache) & vbNewLine
        Msg = Msg & gstrFRODN & vbNewLine
        'Msg = Msg & "Diese Datei kann nicht geöffnet werden." & vbNewLine & vbNewLine
        Msg = Msg & LoadResString(1376 + Sprache) & vbNewLine & vbNewLine
        
        'Msg = Msg & "Entweder die Datei existiert nicht," & vbNewLine & vbNewLine
        Msg = Msg & LoadResString(2208 + Sprache) & vbNewLine & vbNewLine
        
        'Msg = Msg & "oder es ist keine Anwendung mit der" & vbNewLine
        Msg = Msg & LoadResString(1378 + Sprache) & vbNewLine
        'Msg = Msg & "Dateinamen-Erweiterung(Datei-Typ) " & DateinamenErweiterung & " verknüpft." & vbNewLine
        Msg = Msg & LoadResString(1379 + Sprache) & DateinamenErweiterung & LoadResString(1380 + Sprache) & vbNewLine
        'Msg = Msg & "Wählen Sie selbst eine geignete Anwendung, zB mittels Windows-Explorer" & vbNewLine
        Msg = Msg & LoadResString(2012 + Sprache) & vbNewLine
        'Msg = Msg & "Rechtklicken auf den Dateiname -> Öffnen mit... -> Programm auswählen"
        Msg = Msg & LoadResString(2013 + Sprache)
        'MsgBox Msg
        MessageBoxW 0, StrPtr(Msg), StrPtr(LoadResString(1119 + Sprache)), vbInformation
    Else
        'IPTCPresent = False setzen
        frmGridAndThumb.Adodc1.Recordset("IPTCPresent") = 0                                 'Gerbing 13.12.2016
    End If
End Sub

Private Sub mnuVersion_Click()                                                  'Gerbing 01.09.2017
    'me.MousePointer = vbDefault                                                'Gerbing 29.07.2007
    AboutForm.Show 1
    If gblnComefromVideo = True Then                                            'Gerbing 16.06.2012
        frmVideo.Show
    End If
End Sub

Private Sub mnuWeiterselektieren_Click()
    Dim pos As Long
    Dim pos1 As Long
    Dim strLinks As String
    Dim strRechts As String
    Dim SQL As String
    
'    SQL = " SELECT *"
'    SQL = SQL & " FROM " & "Fotos"
'    SQL = SQL & " WHERE Merker<>0
'    SQL = SQL & " ORDER BY Dateiname" & ";"
    SQL = " SELECT *"
    SQL = SQL & " FROM Fotos"
    SQL = SQL & " WHERE " & LoadResString(2524 + Sprache) & "<>0"
    SQL = SQL & " ORDER BY " & LoadResString(1028 + Sprache) & ";"
    With rstsql
        .Source = SQL
        .ActiveConnection = DBado                                                       'Gerbing 23.11.2017
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    If rstsql.EOF Then
        rstsql.Close
        MsgBox LoadResString(3061 + Sprache) 'Es gibt keine mit Merkerspalte markierten Sätze
        Exit Sub
    End If
    '-----------------------------------------------------------------------------------------
    SQL = Query.SQL
    pos = InStr(1, SQL, "WHERE", vbTextCompare)
    pos1 = InStr(1, SQL, "ORDER BY", vbTextCompare)
    strLinks = Left(SQL, pos1 - 1)
    strRechts = Right(SQL, Len(SQL) - pos1 + 1)
    If pos <> 0 Then
        'WHERE gibt es schon
        'AND Merker = 1 hinzufügen
        strLinks = Replace(strLinks, "WHERE", "WHERE (", , , vbTextCompare)         'Gerbing 19.01.2007
        SQL = strLinks & ")"
        'SQL = SQL & " AND " & LoadResString(2524 + Sprache) & "=1 " & strRechts
        SQL = SQL & " AND " & LoadResString(2524 + Sprache) & "<>0 " & strRechts    'Gerbing 26.07.2012
    Else
        'WHERE gibt es noch nicht
        'WHERE Merker<>0 hinzufügen
        SQL = strLinks & " WHERE " & LoadResString(2524 + Sprache) & "<>0 " & strRechts
    End If
    On Error Resume Next
    frmGridAndThumb.rsDataGrid.Close
    On Error GoTo 0
    If gblnSchreibgeschützt = True Then
    ' Recordset erstellen und öffnen
        Set frmGridAndThumb.rsDataGrid = New ADODB.Recordset
        With frmGridAndThumb.rsDataGrid
            .Source = SQL
            .ActiveConnection = DBado                                                       'Gerbing 23.11.2017
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
    Else
        ' Recordset erstellen und öffnen
        Set frmGridAndThumb.rsDataGrid = New ADODB.Recordset
        With frmGridAndThumb.rsDataGrid
            .Source = SQL
            .ActiveConnection = DBado                                                       'Gerbing 23.11.2017
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
    End If
    Set frmGridAndThumb.Adodc1.Recordset = frmGridAndThumb.rsDataGrid
    Set frmGridAndThumb.DBGridNeu.DataSource = frmGridAndThumb.rsDataGrid
    frmGridAndThumb.DBGridNeu.ReBind
    Call frmGridAndThumb.SetSpaltenBreite
End Sub

