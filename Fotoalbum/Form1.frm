VERSION 5.00
Object = "{A8F9B8E7-E699-4FCE-A647-72C877F8E632}#1.8#0"; "editctlsu.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   ClientHeight    =   9288
   ClientLeft      =   192
   ClientTop       =   516
   ClientWidth     =   12336
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   9288
   ScaleWidth      =   12336
   StartUpPosition =   1  'Fenstermitte
   WindowState     =   2  'Maximiert
   Begin VB.Timer TimerKeyboardHook 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5160
      Top             =   1920
   End
   Begin EditCtlsLibUCtl.TextBox txtBildbeschreibung 
      Height          =   372
      Left            =   0
      TabIndex        =   4
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
      CueBanner       =   "Form1.frx":038A
      Text            =   "Form1.frx":03AA
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1332
      Left            =   120
      ScaleHeight     =   1284
      ScaleWidth      =   2004
      TabIndex        =   3
      Top             =   840
      Width           =   2052
      Begin VB.Shape Shape1 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   7  'Diagonalkreuz
         Height          =   972
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   1092
      End
   End
   Begin VB.TextBox txtFont 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3600
      TabIndex        =   1
      Text            =   "txtFont"
      Top             =   480
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.Timer TimerVideoDuration 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6600
      Top             =   1320
   End
   Begin VB.Timer TimerToPlayVideo 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5880
      Top             =   1320
   End
   Begin VB.Timer TimerNachFormLoad 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5160
      Top             =   1320
   End
   Begin VB.Label lblLeereForm 
      Caption         =   "Wählen Sie ein anderes Bild, Tasten F2/F3 oder Pfeil-Tasten"
      Height          =   492
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Visible         =   0   'False
      Width           =   7812
   End
   Begin VB.Label lblVideoWirdgeladen 
      Caption         =   "Video wird geladen. Drücken Sie F5 zur Anzeige des Listenfensters"
      Height          =   492
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Visible         =   0   'False
      Width           =   7812
   End
   Begin VB.Menu mnuDatei 
      Caption         =   "Datei..."
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
         Caption         =   "&Fotosmdb starten"
      End
      Begin VB.Menu mnuRenammdbStarten 
         Caption         =   "&Renammdb starten"
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim Msg As String
    Dim lupe As Integer
    Dim Image1Top As Integer
    Dim Image1Left As Integer
    Dim Image1Width As Integer
    Dim Image1Height As Integer
    Dim Durchläufe As Integer
    Public FotoAlbumTitle As String
    Dim Vergrößerung As Integer
    Public KommentarFensterEinblenden As Boolean
    Dim SQL As String
    Dim BookM As String
    Public StartX As Long
    Public StartY As Long
    Public EndX As Long
    Public EndY As Long
    Public F6Continous As Boolean
    Public SGVH As Double      'Screen-Größen-Breiten-Verhältnis                               'Gerbing 24.09.2009
    Dim BHV As Double       'BreitenHöhenVerhältnis
    Dim BHVImage As Double
    Dim BHVImageVorige As Double
    Dim PIV As Double   'Picture1.Width/Image1.Width-Verhältnis
    Dim RX As Long
    Public blnUnloadExportForm As Boolean
    Public F5Feld1 As String
    Public F5Feld2 As String
    Public F5Feld3 As String
    Public F5Feld4 As String
    Public F5Feld5 As String
    Private adoRs As ADODB.Recordset                                                                'Gerbing 04.01.2006
    Public blnTimer1Enabled As Boolean                                                             'Gerbing 29.03.2012
    Public lngTimer1Interval As Long                                                               'Gerbing 29.03.2012
    

    Private disp(300) As Integer 'the displacement of the HRipples
    Private DisplayHight As Long
    Private DisplaySize As Long
    
    Public EXF As New clsEXIF                                   'Gerbing 07.05.2007
    
    Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer  'Gerbing 07.11.2011
    Private Const KeyPressed As Integer = -32767
    Private Const VK_CONTROL = &H11
    Private Const VK_MENU As Long = &H12&
    Private Const VK_SHIFT As Long = &H10&
    Private Const VK_CAPITAL As Long = &H14&
'    VK_CLEAR = &HC
'    VK_RETURN = &HD
'    VK_SHIFT = &H10
'    VK_CONTROL = &H11
'    VK_MENU = &H12
'    VK_PAUSE = &H13
'    VK_CAPITAL = &H14
'    VK_ESCAPE = &H1B
'    VK_PRIOR = &H21
'    VK_NEXT = &H22
'    VK_HOME = &H24
'    VK_BAK = &H8
'    VK_TAB = &H9
'    VK_LEFT = &H25
'    VK_UP = &H26
'    VK_RIGHT = &H27
'    VK_DOWN = &H28
'    VK_SELECT = &H29
'    VK_END = &H23
'    VK_SNAPSHOT = &H2C
'    VK_INSERT = &H2D
'    VK_DELETE = &H2E
'    VK_HELP = &H2F
'    VK_F1 = &H70
'    VK_F2 = &H71
'    VK_F3 = &H72
'    VK_F4 = &H73
'    VK_F5 = &H74
'    VK_F6 = &H75
'    VK_F7 = &H76
'    VK_F8 = &H77
'    VK_F9 = &H78
'    VK_F10 = &H79
'    VK_F11 = &H7A
'    VK_F12 = &H7B
'    VK_F13 = &H7C
'    VK_F14 = &H7D
'    VK_F15 = &H7E
'    VK_F16 = &H7F
'    VK_NUMLOCK = &H90
'    VK_SCROLL = &H91
'    VK_WIN = &H5B
'    VK_APPS = &H5D
'
    Dim blnHilfeboxStehenLassen As Boolean
    Public lngPointer As Long                                                                               'Gerbing 08.10.2014
    Dim MyZoomPercent As Long
    Public sngWidth As Single
    Public sngHeight As Single
    Public X As Long
    Public Y As Long
    Dim retcode As Long              ' Funktions-Rückgaben
    Public blnComeFromError As Boolean                                                                      'Gerbing 15.02.2014
    Public Enum eGradientDirection                                                                          'Gerbing 15.02.2014
        GRADIENT_FILL_RECT_H = &H0
        GRADIENT_FILL_RECT_V = &H1
    End Enum
    Private SammelTextGetAsyncKeyState As String
    Public GPSLatitude As String                                                                            'Gerbing 27.08.2015
    Public GPSLatitudeRef As String
    Public GPSLongitude As String
    Public GPSLongitudeRef As String
    Dim lngRadius As Long                                                                               'Gerbing 25.04.2018
    
    Private Enum RotateFlipType                                                                         'Gerbing 10.05.2019
        RotateNoneFlipNone = 0
        Rotate90FlipNone = 1
        Rotate180FlipNone = 2
        Rotate270FlipNone = 3
        RotateNoneFlipX = 4
        Rotate90FlipX = 5
        Rotate180FlipX = 6
        Rotate270FlipX = 7
        RotateNoneFlipY = Rotate180FlipX
        Rotate90FlipY = Rotate270FlipX
        Rotate180FlipY = RotateNoneFlipX
        Rotate270FlipY = Rotate90FlipX
        RotateNoneFlipXY = Rotate180FlipNone
        Rotate90FlipXY = Rotate270FlipNone
        Rotate180FlipXY = RotateNoneFlipNone
        Rotate270FlipXY = Rotate90FlipNone
    End Enum
    Private Declare Function GdipImageRotateFlip Lib "gdiplus" _
                            (ByVal Image As Long, ByVal rfType As RotateFlipType) As Status             'Gerbing 10.05.2019
                            
    'Gerbing 04.07.2019
    Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
    Private Const LOCALE_SDECIMAL = &HE                 '  decimal separator



    
Private Sub Form_Initialize()
    Dim returncode As Long
    
    InitCommonControls
    init_global
    Set IniFso = New FileSystemObject
    Set Fso = New FileSystemObject
End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'KeyCode = 0, weil sonst aus unbekanntem Grund Keycode zweimal verarbeitet wird
    'Debug.Print "Form1.Form_KeyDown Keycode=" & eyCode & " Shift=" & Shift
    
    Dim Korrektur As Integer
    Dim DateinamenErweiterung As String
    Dim temp As String
    Dim temp1 As String
    Dim rc As Long
    Dim ShiftDown As Boolean
    Dim AltDown As Boolean
    Dim CtrlDown As Boolean
    Dim n As Long
    Dim start As Long
    Dim Dateiname As String
    Dim mark As Variant
    Dim tempFRODN As String
    Dim i As Long
    
    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
    
    If Shift = vbAltMask And KeyCode = 18 Then  '18 = Menu key                                  'Gerbing 01.09.2017
        KeyCode = 0
        mnuDatei.Visible = Not mnuDatei.Visible
        mnuTools.Visible = Not mnuTools.Visible
        mnuVersion.Visible = Not mnuVersion.Visible
        mnuHilfe.Visible = Not mnuHilfe.Visible
        Sleep 100  'Sonst flackert es bei Festhalten der Taste Alt                              'Gerbing 05.11.2017
        DoEvents                                                                                'Gerbing 20.01.2018
        Exit Sub
    End If
    
    If gblnWeiterMitLeererDatenbank = True Then                                                 'Gerbing 26.01.2006
        If KeyCode = vbKeyI Then
            If gblnSQLServerVersion = True Then
                MsgBox "Diese Funktion wird beim SQL-Server nicht unterstützt"
                MsgBox LoadResString(1829 + Sprache)
            Else
                If Shift = vbCtrlMask Then                              'Strg+I gleichzeitig
                    Call MediaPlayerStop
                    ImportForm.Show 1
                End If
            End If
                
            Exit Sub
        Else
            Exit Sub
        End If
    End If
    
    On Error Resume Next                                                                        'Gerbing 22.08.2007
    'Hilfebx.Hide                                                                               'Gerbing 16.06.2012
    'KommentarForm.Hide                                                                         'Gerbing 16.06.2012
    On Error GoTo 0
    '------------------------------------------------------------------------------------------------
    Select Case KeyCode
        Case vbKeyAdd                                                                           'Gerbing 28.03.2014
            If Shift = vbCtrlMask Then                                  'Strg+'+' gleichzeitig
                GoTo GotovbKeyF4
            End If
        Case vbKeySubtract                                                                      'Gerbing 28.03.2014
            If Shift = vbCtrlMask Then                                  'Strg+'-' gleichzeitig
                GoTo GotovbKeyF1
            End If
        Case vbKeyHome                                                  'Pos1-Taste geht zum Anfang
            Call MediaPlayerStop
            KeyCode = 0
            Call GeheZumAnfang
            Call BildAnzeigen
        '--------------------
        Case 33                                                         'Bild nach oben geht zum Anfang
            Call MediaPlayerStop
            KeyCode = 0
            Call GeheZumAnfang
            Call BildAnzeigen
        Case 34                                                         'Bild nach unten geht zum Ende
            Call MediaPlayerStop
            KeyCode = 0
            Call GeheZumEnde
            Call BildAnzeigen
        '--------------------
        Case vbKeyEnd                                                   'Ende-Taste geht zum Ende
            Call MediaPlayerStop
            KeyCode = 0
            Call GeheZumEnde
            Call BildAnzeigen
        '--------------------
        Case vbKeyF1                                                    'F1
            If blnComeFromError = True Then                                                     'Gerbing 15.02.2014
                MakeGradient Me, vbBlue, vbGreen, GRADIENT_FILL_RECT_V
                MakeGradient Form1.Picture1, vbBlue, vbGreen, GRADIENT_FILL_RECT_V
                'MsgBox gstrFRODN & " Bild kann nicht geladen werden"
                Msg = gstrFRODN & " " & LoadResString(2056 + Sprache)
                MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
                Exit Sub                                                                        'Gerbing 15.02.2014
            End If
GotovbKeyF1:
            If gblnComefromVideo = True Then
            'If frmVideo.WMP.URL <> "" Then                                                     'Gerbing 01.09.2008
                'Bei Videos soll F1 wirkungslos sein
                'frmVideo.Show                                                                  'Gerbing 16.06.2012
                Exit Sub
            End If
            Call MediaPlayerStop
            KeyCode = 0
            glngDiffX = 0
            glngDiffY = 0
            gblnRechteckLupeScharf = False
            If glngZoomProzent < 2 Then
                'MsgBox "Minimumzoom erreicht"
            Else
                glngZoomProzent = glngZoomProzent \ 2
            End If
            Call SpeichernInBildPosList
            Call Form1.MyDrawImage(gstrFRODN, glngZoomProzent)                                  'Gerbing 10.06.2012
        Case vbKeyF2                                                    'F2
            gblnComeFromF2F3 = True                                                             'Gerbing 27.03.2014
            If gblnSQLServerVersion = True Then                                                 'Gerbing 29.12.2011
                Call KontrolleManagement
            Else
'                If gblnVollversion = False Then                                                 'Gerbing 24.06.2015
'                    If gintDiffTage > 365 Then
'                        Copy.Show 1
'                    End If
'                End If
            End If
            Call MediaPlayerStop
            KeyCode = 0
            frmGridAndThumb.Hide                                                                'Gerbing 29.03.2015 sonst flackert es
            Call GeheEinBildZurück
            Call BildAnzeigen
        Case vbKeyLeft                                                  'Alt + Pfeil nach links wirkt wie F2
            If Shift = vbAltMask Then
                gblnComeFromF2F3 = True                                                         'Gerbing 27.03.2014
                If gblnSQLServerVersion = True Then                                             'Gerbing 29.12.2011
                    Call KontrolleManagement
                Else
'                    If gblnVollversion = False Then                                             'Gerbing 24.06.2015
'                        If gintDiffTage > 365 Then
'                            Copy.Show 1
'                        End If
'                    End If
                End If
                Call MediaPlayerStop
                KeyCode = 0
                frmGridAndThumb.Hide                                                            'Gerbing 29.03.2015 sonst flackert es
                Call GeheEinBildZurück
                Call BildAnzeigen                                                               'Gerbing 10.06.2012
            End If
            If Shift = vbShiftMask Then                                 'Pfeil nach links und gleichzeitig Shift-Taste
                glngDiffX = glngDiffX - screenWidth / 4
                Call SpeichernInBildPosList
                gblnRechteckLupeScharf = False
                KeyCode = 0
                Call Form1.MyDrawImage(gstrFRODN, glngZoomProzent)                              'Gerbing 10.06.2012
            End If
        Case vbKeyF3                                                    'F3
            gblnComeFromF2F3 = True                                                             'Gerbing 27.03.2014
            If gblnSQLServerVersion = True Then                                                 'Gerbing 29.12.2011
                Call KontrolleManagement
            Else
'                If gblnVollversion = False Then                                                 'Gerbing 24.06.2015
'                    If gintDiffTage > 365 Then
'                        Copy.Show 1
'                    End If
'                End If
            End If
            Call MediaPlayerStop
            KeyCode = 0
            frmGridAndThumb.Hide                                                                'Gerbing 29.03.2015 sonst flackert es
            Call GeheEinBildVorwärts
            Call BildAnzeigen
        Case vbKeyRight                                                 'Alt + Pfeil nach rechts wirkt wie F3
            If Shift = vbAltMask Then
                gblnComeFromF2F3 = True                                                         'Gerbing 27.03.2014
                If gblnSQLServerVersion = True Then                                             'Gerbing 29.12.2011
                    Call KontrolleManagement
                Else
'                    If gblnVollversion = False Then                                             'Gerbing 24.06.2015
'                        If gintDiffTage > 365 Then
'                            Copy.Show 1
'                        End If
'                    End If
                End If
                Call MediaPlayerStop
                KeyCode = 0
                frmGridAndThumb.Hide                                                            'Gerbing 29.03.2015 sonst flackert es
                Call GeheEinBildVorwärts
                Call BildAnzeigen
            End If
            If Shift = vbShiftMask Then                                 'Pfeil nach rechts und gleichzeitig Shift-Taste
                glngDiffX = glngDiffX + screenWidth / 4
                Call SpeichernInBildPosList
                gblnRechteckLupeScharf = False
                KeyCode = 0
                Call Form1.MyDrawImage(gstrFRODN, glngZoomProzent)                              'Gerbing 10.06.2012
            End If
        Case vbKeyF4                                                    'F4
            If Shift = vbAltMask Then                                   'Alt + F4               'Gerbing 22.05.2013
                Unload frmGridAndThumb
                Unload Hilfebx
                Unload KommentarForm
                Unload Query
                'Unload QueryJedesFeld
                Unload MP
                End
            End If
GotovbKeyF4:
            If blnComeFromError = True Then                                                     'Gerbing 15.02.2014
                MakeGradient Me, vbBlue, vbGreen, GRADIENT_FILL_RECT_V
                MakeGradient Form1.Picture1, vbBlue, vbGreen, GRADIENT_FILL_RECT_V
                'MsgBox gstrFRODN & " Bild kann nicht geladen werden"
                Msg = gstrFRODN & " " & LoadResString(2056 + Sprache)
                MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
                Exit Sub                                                                        'Gerbing 15.02.2014
            End If
            If gblnComefromVideo = True Then
            'If frmVideo.WMP.URL <> "" Then                                                     'Gerbing 01.09.2008
                'Bei Videos soll F4 wirkungslos sein
                'frmVideo.Show                                                                  'Gerbing 16.06.2012
                Exit Sub
            End If
            Call MediaPlayerStop
            KeyCode = 0
            glngDiffX = 0
            glngDiffY = 0
            gblnRechteckLupeScharf = False
            On Error Resume Next                                                                'Gerbing 05.09.2012
            glngZoomProzent = glngZoomProzent * 2
            On Error GoTo 0                                                                     'Gerbing 05.09.2012
            Call SpeichernInBildPosList
            Call Form1.MyDrawImage(gstrFRODN, glngZoomProzent)                                  'Gerbing 10.06.2012
        Case vbKeyF5                                                    'F5
            If gblnSQLServerVersion = True Then                                                 'Gerbing 29.12.2011
                Call KontrolleManagement
            Else
                If gblnVollversion = False Then                                                 'Gerbing 13.10.2005
                    If gintDiffTage > 180 Then
                        Copy.Show 1
                    End If
                End If
            End If
            
            'Call MediaPlayerStop
            'KeyCode = 0        'Gerbing 16.06.2012   'nicht löschen, sonst wird Shift nicht erkannt, wenn ich aus frmVideo komme
            'Für einen eventuellen Doppel-Click muß der Mauszeiger sichtbar sein
            'Me.MousePointer = vbDefault                                                        'Gerbing 29.07.2007
            If gstrCommandLine <> "/WRITE" Then
                frmGridAndThumb.DBGridNeu.AllowUpdate = True                                         'Gerbing 21.11.2007
                frmGridAndThumb.DBGridNeu.Columns(0).Locked = False      'in die Merker-Spalte darf man schreiben
                For n = 1 To frmGridAndThumb.DBGridNeu.Columns.Count - 1
                    frmGridAndThumb.DBGridNeu.Columns(n).Locked = True
                Next n
            Else
                frmGridAndThumb.DBGridNeu.AllowUpdate = True
            End If
            If gblnSchreibgeschützt = True Then                                                     'Gerbing 04.01.2006
                frmGridAndThumb.DBGridNeu.AllowUpdate = False
            End If
            '-----------------------------------------------------------
            'If Shift = vbCtrlMask Then
            If Shift = vbShiftMask Then                                 'F5 und Taste Umsch gleichzeitig            soll F5MehrereZeilen öffnen
                'Berücksichtigung der nutzerdefinierten Felder                                      'Gerbing 14.02.2005
                'Call FelderAusfüllenF5MehrereZeilen
                Unload F5MehrereZeilen                                                              'Gerbing 07.03.2016
                F5MehrereZeilen.Show 1                                                              'Gerbing 06.11.2006
                If gblnComefromVideo = True Then                                                    'Gerbing 16.06.2012
                    frmVideo.Show
                End If
            Else
                'F5                                                     'F5
                If gblnComefromVideo = True Then                                                    'Gerbing 16.06.2012
                    frmVideo.Show
                End If
                frmGridAndThumb.WindowState = vbNormal                                              'Gerbing 04.05.2015
                frmGridAndThumb.Show                                                                'Gerbing 15.11.2012
                On Error Resume Next
                If frmGridAndThumb.DBGridNeu.SelBookmarks.Count = 1 Then                            'Gerbing 30.11.2012
                    frmGridAndThumb.DBGridNeu.SelBookmarks.Remove 0                                 'Gerbing 30.11.2012
                End If                                                                              'Gerbing 30.11.2012
                frmGridAndThumb.DBGridNeu.SelBookmarks.Add frmGridAndThumb.rsDataGrid.Bookmark      'Gerbing 30.11.2012
                On Error GoTo 0
'                If Query.chkFensterGrößeÄnderbar.Value = 1 Then                                    'Gerbing 06.12.2005 29.03.2015
'                    frmGridAndThumb.Width = Form1.Width
'                Else
'                    frmGridAndThumb.Width = Screen.Width
'                End If
                If Shift = vbAltMask Then                               'F5 und Taste Alt gleichzeitig     'Gerbing 22.04.2014
                    'F5 + Alt soll den Schalter blnF5Alt einschalten
                    gblnF5Alt = True
                Else
                    gblnF5Alt = False
                End If
                If gblnF5Alt = True Then                                                   'Gerbing 22.04.2014
                    frmGridAndThumb.DBGridNeu.Top = 0
                    frmGridAndThumb.DBGridNeu.height = frmGridAndThumb.DBGridNeu.height + 80        'Gerbing 29.03.2015
                    frmGridAndThumb.btnSpaltenbreitenSpeichern.Visible = False
                    frmGridAndThumb.btnMerkerspalteEinschalten.Visible = False
                    frmGridAndThumb.btnRefresh.Visible = False
                    frmGridAndThumb.DBGridNeu.FirstRow = frmGridAndThumb.DBGridNeu.Bookmark           'Gerbing 10.09.2014                'DBGridNeu.Row muss sichtbar sein, sonst
                    If frmGridAndThumb.DBGridNeu.SelBookmarks(i) - frmGridAndThumb.DBGridNeu.FirstRow >= 0 Then
                        frmGridAndThumb.DBGridNeu.Row = frmGridAndThumb.DBGridNeu.SelBookmarks(i) - frmGridAndThumb.DBGridNeu.FirstRow      'Laufzeitfehler 6148 Ungültige Zeilennummer
                    End If
                Else
                    frmGridAndThumb.DBGridNeu.Top = frmGridAndThumb.btnSpaltenbreitenSpeichern.height + frmGridAndThumb.btnSpaltenbreitenSpeichern.height + 30 'Gerbing 29.03.2015
                    frmGridAndThumb.btnSpaltenbreitenSpeichern.Visible = True
                    frmGridAndThumb.btnMerkerspalteEinschalten.Visible = True
                    frmGridAndThumb.btnRefresh.Visible = True
                    If Query.optNurErstenTreffer.Value = True Or Query.optErsterZufallstreffer.Value = True Then        'Gerbing 09.02.2013
                        frmGridAndThumb.btnMerkerspalteEinschalten.Visible = False                   'Gerbing 23.01.2007
                    Else
                        frmGridAndThumb.btnMerkerspalteEinschalten.Visible = True
                    End If
                End If
                Call AnpassenNutzerWunsch(Me)                                                   'Gerbing 11.03.2017
                Call AnpassenHeadFont(frmGridAndThumb.DBGridNeu)                                'Gerbing 23.06.2011
                frmGridAndThumb.DBGridNeu.ZOrder
                frmGridAndThumb.DBGridNeu.Refresh
                'ich brauche ein Mittel um frmGridAndThumb.DbGridNeu zu aktivieren, die aktuelle Zeile soll den Fokus
                'besitzen, damit ich dort sofort etwas hineinschreiben kann
'                temp = frmGridAndThumb.DBGridNeu.Columns(1)
'                frmGridAndThumb.DBGridNeu.Columns(1) = temp
                'frmGridAndThumb.DBGridNeu.SetFocus                                              Gerbing 15.11.2012
                Call frmGridAndThumb.SetSpaltenBreite
                KeyCode = 0                                                                 'Gerbing 15.11.2012
            End If
        Case vbKeyF6                                                    'F6
            If Shift = vbCtrlMask Then                                  'Gerbing 03.11.2004 Strg+F6 gleichzeitig
                On Error Resume Next
                Call FRODateiname
                DateinamenErweiterung = Right(gstrFRODN, 3)
                If Err = 91 Then    'Objektvariable oder With-Blockvariable nicht festgelegt
                    Msg = "Es wurde kein einziger Datensatz gefunden." & NL
                    Msg = Msg & "Mit der F8-Taste können Sie die Suche wiederholen"
                    'MsgBox msg                 'Gerbing 08.11.2005
                    MsgBox LoadResString(2007 + Sprache) & NL & LoadResString(2008 + Sprache)
                    Exit Sub
                End If
                DateinamenErweiterung = UCase(DateinamenErweiterung)
                On Error GoTo 0
                Form1.blnTimer1Enabled = True                                               'Gerbing 29.03.2012
                EnableTimer Form1.lngTimer1Interval
                Form1.F6Continous = True
                Select Case DateinamenErweiterung
                    Case "AVI", "MPG", "PEG", "MOV", "MPE", "ASF", "ASX", "WMV", "MP4", "MKV", "FLV"      'Gerbing 10.12.2017
                        If StrComp(Right(gstrFRODN, 4), "JPEG", vbTextCompare) = 0 Then     'Gerbing 21.08.2006
                            'nichts tun
                        Else
                            'damit bei videos nicht alle 3 Sekunden der Timer drankommt
                            'sondern erst wenn das video aufgehört hat
                            'Diese Steuerung übernimmt Form1.F6Continous
                            Form1.blnTimer1Enabled = False
                            DisableTimer
                        End If
                End Select
            End If
        Case vbKeyF7                                                    'F7
            Form1.F6Continous = False
            Form1.blnTimer1Enabled = False
            DisableTimer                                                                    'Gerbing 29.03.2012
        Case vbKeyF8                                                    'F8
'            If frmGridAndThumb.blnComeFromBtnMitThumbnailsClick = True Then
'                Call Query.Beenden                                                          'Gerbing 09.06.2015
'            End If
            If lngPointer Then                                                              'Gerbing 31.12.2012 und 04.09.2013
                retcode = GdipDisposeImage(lngPointer)
                lngPointer = 0                                                              'Gerbing 19.04.2017
            End If
            If m_lngGraphics Then                                                           'Gerbing 31.12.2012 und 04.09.2013
                If GdipDeleteGraphics(m_lngGraphics) Then _
                    'MsgBox "Graphics object could not be deleted", vbCritical
                End If
            End If
            gblnFotosMitFET = False                                                         'Gerbing 16.02.2013
            Query.optAlleTreffer.Value = True                                               'Gerbing 16.02.2013
            Form1.Hide                                                                      'Gerbing 10.06.2012
            frmVideo.Hide                                                                   'Gerbing 16.06.2012
            SQLWurdeBearbeitet = False
            KommentarFensterEinblenden = False
            Unload KommentarForm                                                            'Gerbing 23.10.2012
'            Unload frmGridAndThumb                                                          'Gerbing 29.03.2015
            Query.OKGewählt = False                                                         'Gerbing 06.12.2005
            Me.MousePointer = vbDefault                                                     'Gerbing 29.07.2007
            Call MediaPlayerStop
            frmVideo.WMP.url = ""                                                           'Gerbing 01.09.2008
            KeyCode = 0
            blnTimer1Enabled = False
            DisableTimer                                                                    'Gerbing 29.03.2012
            F6Continous = False                                                             'Gerbing 15.11.2011
            Query.SQLText.Visible = False
            Query.SucheInJedemFeld = False
            Me.Hide
            gblnComeFromButtonF8 = True
            Unload frmGridAndThumb
            gblnComeFromButtonF8 = False
            XYPos.Hide                                                                      'Gerbing 17.04.2005
            Query.Show                             'Suchkriterien festlegen                 'Gerbing 20.06.2012
            If gblnSQLServerVersion = True Then                                             'Gerbing 29.12.2011
                Call KontrolleManagement
            Else
                If gblnVollversion = False Then                                             'Gerbing 13.10.2005
                    If gintDiffTage > 90 Then
                        Copy.Show 1  'muss nach Query.show kommen                           'Gerbing 24.06.2015
                    End If
                End If
            End If
            gblnComeFromF8 = True                                                           'Gerbing 20.06.2012
            gstrRowColChangeName = ""                                                       'Gerbing 06.06.2015
            gblnWasHeadClick = False                                                        'Gerbing 12.12.2016
        Case vbKeyF9                                                    'F9
            If frmVideo.WMP.url = "" Then                                                   'Gerbing 01.09.2008
                'Bei Videos kein MediaPlayerStop
                Call MediaPlayerStop
            End If
            KeyCode = 0
            If gblnMouseSichtbar = True Then
                gblnMouseSichtbar = False
                Me.MousePointer = vbCustom      '99                                         'Gerbing 29.07.2007
                'Me.MouseIcon = LoadPicture(AppPath & "\MOUSE01.ICO")                       'Gerbing 29.07.2007
                Me.MouseIcon = LoadResPicture(104, 1)                                       'Gerbing 04.03.2013
            Else
                gblnMouseSichtbar = True
                If gblnRechteckLupeScharf = True Then                                       'Gerbing 29.07.2007
                    Me.MousePointer = vbCustom
                    'Me.MouseIcon = LoadPicture(AppPath & "\SquareZoom.ico")                'Gerbing 29.07.2007
                    Me.MouseIcon = LoadResPicture(105, 1)                                   'Gerbing 04.03.2013
                Else
                    Me.MousePointer = vbDefault                                             'Gerbing 29.07.2007
                End If
            End If
        Case vbKeyF10                                                   'F10
            KommentarFensterEinblenden = True
            'Gerbing 22.07.2005
            'Kommentar-Fenster stets einblenden, wenn PF10 gedrückt wird                    Gerbing 14.04.2006
            If Not IsNull(frmGridAndThumb.Adodc1.Recordset(LoadResString(1030 + Sprache))) Then
                mark = 1                                                                                                        'Gerbing 27.10.2012
                tempFRODN = Replace(gstrFRODN, gstrFotosMdbLocation & "\", "+:\")                                               'Gerbing 27.10.2012
                tempFRODN = Replace(tempFRODN, "'", "''")                                   'Gerbing 23.01.2018
                'Bei Dateinamen mit Hochkomma bringt ..Find Laufzeitfehler -> ersetzen durch 2 Hochkommas  'Gerbing 23.01.2018
                'frmGridAndThumb.rsDataGrid.Find LoadResString(1028 + Sprache) & " = '" & tempFRODN & "'"                            'Gerbing 27.10.2012
                frmGridAndThumb.rsDataGrid.Find LoadResString(1028 + Sprache) & " = '" & tempFRODN & "'", 0, adSearchForward, mark   'Gerbing 27.10.2012
                
                If frmGridAndThumb.rsDataGrid.EOF Then
                    Msg = tempFRODN & vbNewLine
                    Msg = Msg & "frmGridAndThumb.rsDataGrid.Find"
                    'MsgBox Msg
                    MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
                    Exit Sub
                End If
                
                KommentarForm.txtKommentar.Text = frmGridAndThumb.rsDataGrid(LoadResString(1030 + Sprache))
            Else
                KommentarForm.txtKommentar.Text = ""                                        'Gerbing 30.05.2006
            End If
            KommentarForm.Show
'            KommentarForm.SetFocus                                                          'Gerbing 16.04.2006 11.09.2014
'            On Error Resume Next
'            AppActivate FotoAlbumTitle
'            On Error GoTo 0
        Case vbKeyF11                                                   'F11
            KommentarFensterEinblenden = False
            If gblnComefromVideo = False Then
            End If
            'KommentarForm.Hide                                                             'Gerbing 30.12.2005
            Unload KommentarForm                                                            'Gerbing 23.10.2012
        Case vbKeyF12                                                   'F12
            'me.MousePointer = vbDefault                                                    'Gerbing 29.07.2007
            'Call AnpassenFontSize(WertxForm)                                               'Gerbing 23.06.2011
            KeyCode = 0                                                                     'Gerbing 16.06.2012
            WertxForm.Show 1                                                                'Gerbing 16.06.2012
            'Form1.Show                                                                     'Gerbing 05.08.2013
'            If gblnComefromVideo = False Then                                               'Gerbing 23.10.2013 04.10.2019
'                frmVideo.lblLeereForm.Visible = True
'            End If
        Case vbKeyB
            If frmVideo.WMP.url = "" Then                                                   'Gerbing 01.09.2008
                'Bei Videos keine XY Position
                If Shift = vbCtrlMask Then                              'Strg+B gleichzeitig
                    'me.MousePointer = vbDefault                                            'Gerbing 29.07.2007
                    XYPos.WindowState = vbNormal                                            'Gerbing 04.05.2015
                    If Query.chkFensterGrößeÄnderbar.Value = 1 Then                         'Gerbing 06.12.2005
                        XYPos.Top = Form1.Top                                               'Gerbing 06.12.2006
                        XYPos.Left = Form1.Left
                    End If
                    XYPos.Show
                End If
            End If
        Case vbKeyZ
            If gblnRechteckLupeScharf = True Then Exit Sub                                  'Gerbing 23.10.2014
            If Shift = vbCtrlMask Then                                  'Strg+Z gleichzeitig
                If blnComeFromError = True Then                                             'Gerbing 15.02.2014
                    MakeGradient Me, vbBlue, vbGreen, GRADIENT_FILL_RECT_V
                    MakeGradient Form1.Picture1, vbBlue, vbGreen, GRADIENT_FILL_RECT_V
                    'MsgBox gstrFRODN & " Bild kann nicht geladen werden"
                    Msg = gstrFRODN & " " & LoadResString(2056 + Sprache)
                    MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
                    Exit Sub                                                                'Gerbing 15.02.2014
                End If
                Call BildAnzeigen                                                           'Gerbing 23.10.2014
                gblnRechteckLupeScharf = True                                               'Gerbing 23.10.2014 zuerst das Bild neu zeichnen
                Me.MousePointer = vbCustom                                                  'Gerbing 29.07.2007
                'Me.MouseIcon = LoadPicture(AppPath & "\SquareZoom.ico")                    'Gerbing 29.07.2007
                Me.MouseIcon = LoadResPicture(105, 1)                                       'Gerbing 04.03.2013
            End If
        Case vbKeyO
            If Shift = vbCtrlMask Then                                  'Strg+O gleichzeitig
                gblnRechteckLupeScharf = False
                Me.MousePointer = vbDefault                                                 'Gerbing 29.07.2007
                Call BildAnzeigen                                                           'Gerbing 20.02.2013
            End If
        Case 38                                                         'Pfeil nach oben
            If Shift = vbShiftMask Then                                 'Pfeil nach oben und gleichzeitig Shift-Taste
                glngDiffY = glngDiffY - screenWidth / 4
                Call SpeichernInBildPosList
                gblnRechteckLupeScharf = False
                KeyCode = 0
                Call Form1.MyDrawImage(gstrFRODN, glngZoomProzent)                          'Gerbing 10.06.2012
            End If
        Case 40                                                         'Pfeil nach unten
            If Shift = vbShiftMask Then                                 'Pfeil nach unten und gleichzeitig Shift-Taste
                glngDiffY = glngDiffY + screenWidth / 4
                Call SpeichernInBildPosList
                gblnRechteckLupeScharf = False
                KeyCode = 0
                Call Form1.MyDrawImage(gstrFRODN, glngZoomProzent)                          'Gerbing 10.06.2012
            End If
        Case vbKeyI
            If Shift = vbCtrlMask Then                                  'Strg+I gleichzeitig
                If gblnSQLServerVersion = True Then
                    MsgBox "Diese Funktion wird beim SQL-Server nicht unterstützt"
                    MsgBox LoadResString(1829 + Sprache)
                Else
                    Call MediaPlayerStop
                    ImportForm.Show 1
                End If
            End If
            frmVideo.lblLeereForm.Visible = True                                            'Gerbing 04.12.2012
        Case vbKeyK
            If Shift = vbCtrlMask Then                                  'Strg+K gleichzeitig
                Call MediaPlayerStop
                KeyCode = 0
                'me.MousePointer = vbDefault                                                'Gerbing 29.07.2007
                blnUnloadExportForm = False
                ExportForm.Show 1
            End If
            frmVideo.lblLeereForm.Visible = True                                            'Gerbing 04.12.2012
        Case vbKeyN
            If CtrlDown And ShiftDown Then                              'Strg+Num+N gleichzeitig                  'Gerbing 03.03.2012
                gblnBildBeschreibung = True
                If gblnComefromVideo = True Then                                            'Gerbing 28.08.2020
                    Call VideoAbspielen                                                     'Gerbing 28.08.2020
                Else
                    Call MyDrawImage(gstrFRODN, glngZoomProzent)                            'Gerbing 28.08.2020
                End If
            End If
        Case vbKeyM
            If CtrlDown And ShiftDown Then                              'Strg+Num+M gleichzeitig                  'Gerbing 03.03.2012
                gblnBildBeschreibung = False
                If gblnComefromVideo = True Then                                            'Gerbing 28.08.2020
                    Call VideoAbspielen                                                     'Gerbing 28.08.2020
                Else
                    Call MyDrawImage(gstrFRODN, glngZoomProzent)                            'Gerbing 28.08.2020
                End If
            End If
        Case vbKeyG
            If Shift = vbCtrlMask Then                                  'Strg+G gleichzeitig                    'Gerbing 03.10.2016
                If Not (gblnVollversion = True And gblnProversion = True) Then              'Gerbing 27.09.2016 02.10.2019
                    Msg = LoadResString(2335 + Sprache) 'Für diese Funktion benötigen Sie die Professional Version.
                    MsgBox Msg
                    Exit Sub
                End If
                Call ZeigeGEOPosition
            End If
        Case vbKeyC                                                     'Strg+C gleichzeitig Gerbing 11.08.2017
            If Shift = vbCtrlMask Then
                frmZwischenablageOderOrdner.Show 1
            End If
        Case vbKeySeparator                                             'vbKeySeparator
            'vbKeySeparator gilt als Verschiebung mit der Maus
            'vbKeySeparator gilt als Kennzeichnung für Rechtecklupe
            'vbKeySeparator gilt als als Kennzeichnung für Neuzeichnen anstelle Zoom
            'vbKeySeparator (Enter auf dem numerischen Block)
            Call Form1.MyDrawImage(gstrFRODN, glngZoomProzent)                              'Gerbing 10.06.2012
        Case Else
            'sämtliche anderen Tasten drücken würde sonst das Video verschwinden lassen
            If gblnComefromVideo = True Then                                                'Gerbing 16.06.2012
                frmVideo.Show
            End If
    End Select
End Sub

Private Sub Form_Load()
'   Form1.Borderstyle = 2 (änderbar)
'   Form1.Caption wird erst nach Form_Load in TimerNachFormLoad_Timer mit einem String gefüllt
'   Form1.Controlbox = True read-only at run time                   'Gerbing 20.11.2008
'   Form1.Picture1.AutoRedraw = True                                'Gerbing 23.10.2014 sonst bleiben schwarze Flecken
'   Form1.Picture1.AutoSize = False

    Dim Msg As String                                               'Gerbing 08.03.2020
    Dim sSource As String                                           'Gerbing 23.11.2017
    Dim sDest As String
    Dim Datei As String
    Dim j As Integer
    Dim DemoDat As Date
    Dim DateiNummer As Long
    Dim diff As Integer
    Dim strFreischalt As String
    Dim Copyright As String
    Dim rc As Long
    Dim WndHnd As Long
    Dim CurrWnd As Long
    Dim TitelLength As Long
    Dim FensterTitel As String
    Dim X As Long
    Dim rstNeu As ADODB.Recordset
    Dim fs As New Scripting.FileSystemObject
    Dim f
    Dim AppId                                                       'Gerbing 08.08.2020

    Const SW_RESTORE = 9
    Const GW_HWNDFIRST = 0
    Const GW_HWNDNEXT = 2
    
    gblnProversion = True                                           'Gerbing 21.11.2019
    gblnVollversion = True                                          'Gerbing 21.11.2019
    Form1.Picture1.AutoRedraw = True                                'Gerbing 23.10.2014 sonst bleiben schwarze Flecken
    Form1.Picture1.AutoSize = False

    gblnErsterStart = True                                          'Gerbing 29.10.2013
    'AppPath = App.Path
    AppPath = getCurrentDir                                         'Gerbing 04.03.2013
    If Right(AppPath, 1) = "\" Then
        AppPath = Left(AppPath, Len(AppPath) - 1)
    End If
    'MessageBoxW 0, StrPtr(AppPath), StrPtr(""), vbInformation
    gstrFotosMdbLocation = AppPath                                  'Gerbing 07.11.2011
    
    'Bei der ersten Benutzung fragen, ob es fotos.mdb gibt          'Gerbing 04.05.2018
    If file_path_exist(AppPath & "\fotos.mdb") = False Then         'Gerbing 04.05.2018
        'fotos.mdb gibt es nicht
        'bei nein fotosStart.mdb umnennen in fotos.mdb
        If file_path_exist(AppPath & "\fotosStart.mdb") = False Then
            'fotosstart.mdb gibt es nicht
            Copyright = Right(FileInfo(AppPath & "\fotos.exe").LegalCopyright, 2) 'Gerbing 18.05.2018
            If Copyright <> "-1" Then                               'Gerbing 18.05.2018
                'Abbruch, wenn es beide Dateien nicht gibt und wenn es nicht die SQL Server Version=Professional Version ist
                Msg = "Both files not found. " & AppPath & "\fotosStart.mdb" & " and " & AppPath & "\fotos.mdb"
                MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
                End
            End If
        End If
        'Umnennen
        'rc = NameAs(Quellname, Zielname)
        rc = NameAs(AppPath & "\fotosStart.mdb", AppPath & "\fotos.mdb")
    Else
        'fotos.mdb gibt es
        'Bei Programmstart                                          'Gerbing 28.05.2019
        '1.Löschen fotos_copy.mdb
        '2.fotos.mdb kopieren in fotos_copy.mdb
        rc = file_delete(AppPath & "\fotos_copy.mdb", , True)
        rc = file_copy(AppPath & "\fotos.mdb", AppPath & "\fotos_copy.mdb")
    End If                                                          'Gerbing 04.05.2018
    
    glngForm1Hwnd = Me.hWnd                                         'Gerbing 09.05.2012
    On Error Resume Next
    'gdtDatumFotosMdb wird nur bei der Shareware-Version gebraucht, damit msdmo.log erzeugt werden kann
    'gdtDatumFotosMdb = FileDateTime(App.Path & "\fotos.mdb")        'Gerbing 05.03.2012
    Set f = fs.GetFile(AppPath & "\fotos.mdb")
    gdtDatumFotosMdb = f.DateLastModified
    On Error GoTo 0
    Call ReadFotosIniFile                                           'Gerbing 04.12.2011
    WriteCDS "0"  'CompactDatabaseStarted=0                         'Gerbing 08.03.2020
    WriteCDE "0"  'CompactDatabaseEnded=0                           'Gerbing 08.03.2020
'--------------------------------------------------------------------------------------------------------------------
    lngTimer1Interval = PublicAutomaticInterval * 1000              'Gerbing 29.03.2012
    
    Call AnpassenNutzerWunsch(Me)                                   'Gerbing 11.03.2017
'    Me.Top = 0                                                      'Gerbing 20.11.2008
'    Me.Left = 0
'    screenWidth = GetDeviceCaps(Me.hDC, HORZRES)                    'Gerbing 29.03.2012
'    screenHeight = GetDeviceCaps(Me.hDC, VERTRES)                   'Gerbing 29.03.2012
'    SGVH = screenWidth / screenHeight                               'Gerbing 24.09.2009
'    'Me.Width = ScreenWidth / 2 * Screen.TwipsPerPixelX             'Gerbing 09.05.2012
'    Me.height = screenHeight * Screen.TwipsPerPixelY
    '----------------------------------------------------------------------------------Gerbing 03.05.2020
    Call SpracheFestlegen           'hier wird auch ermittelt ob es SQL Server Version ist
    '----------------------------------------------------------------------------------
    If gblnSQLServerVersion = False Then                            'Gerbing 03.05.2020
        If (GetAsyncKeyState(VK_SHIFT) = KeyPressed) Then               'Gerbing 07.11.2011
            frmFotoAlbumWirdGeladen.Show                                'Gerbing 27.08.2017
            DoEvents                                                    'Gerbing 27.08.2017
            'FremdeFotosMdb nicht komprimieren                          'Gerbing 27.08.2017
            'FremdeFotosMdb.CompactDatabase dauert 20 Sekunden          'Gerbing 27.08.2017
            Call FremdeFotosMdb                                         'Gerbing 07.11.2011
            '----------------------------------------------------------------------------------Gerbing 03.05.2020
            Call SpracheFestlegen           'wiederholen hier wird auch ermittelt ob es SQL Server Version ist
            '----------------------------------------------------------------------------------
        Else
            'Man sollte die Datenbank gleich beim Start Komprimieren, weil beim Prüfen3             'Gerbing 22.08.2008
            'und bei Arbeit mit 'Nur den ersten Treffer pro Jahr erlauben'
            'der Umfang immer größer wird
            'DBEngine.CompactDatabase gstrFotosMdbLocation & "\fotos.mdb", gstrFotosMdbLocation & "\Newfotos.mdb" 'Gerbing 23.11.2017
            sSource = gstrFotosMdbLocation & "\fotos.mdb"
            sDest = gstrFotosMdbLocation & "\Newfotos.mdb"
            rc = file_delete(gstrFotosMdbLocation & "\Newfotos.mdb", , True)
            '--------------------------------------------------------------------------------------------------
            'Zur Kontrolle, ob CompactDB ausgeführt werden konnte wird vorher CheckCompactDatabase.exe gestartet Gerbing 08.03.2020
            'Wenn CheckCompactDatabase.exe nach 10 Sekunden feststellt, dass in der fotos.ini
            'CompactDatabaseEnded <> 1 ist, dann muss CheckCompactDatabase.exe die Meldung bringen,
            'dass AccessDatabaseEngine.exe wieder holt werden muss und vom Nutzer auch gleich gestartet werden kann
            
            'If Dir(AppPath & "\CheckCompactDatabase.exe") = "" Then
            If file_path_exist(AppPath & "\CheckCompactDatabase.exe") = False Then
                'msg = "CheckCompactDatabase konnte nicht gestartet werden." & vbNewLine
                Msg = LoadResString(2565 + Sprache) & vbNewLine
                'msg = msg & "CheckCompactDatabase.exe muss im gleichen Ordner stehen wie fotos.exe"
                Msg = Msg & LoadResString(2566 + Sprache)
                'MsgBox Msg
                MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
                End
            End If
            AppId = Shell(AppPath & "\CheckCompactDatabase.exe", vbNormalFocus)                                 'Gerbing 08.03.2020
            WriteCDS "1"  'CompactDatabaseStarted=1                                                             'Gerbing 08.03.2020
            '--------------------------------------------------------------------------------------------------
            If CompactDB(sSource, sDest) Then
                'MsgBox "Compact complete"
                If file_path_exist(gstrFotosMdbLocation & "\Newfotos.mdb") = True Then
                    rc = file_delete(gstrFotosMdbLocation & "\fotos.mdb", , True)
                    'rc = file_copy(Quellname, Zielname)                                             'Gerbing 18.10.2017
                    rc = file_copy(gstrFotosMdbLocation & "\Newfotos.mdb", gstrFotosMdbLocation & "\fotos.mdb") 'Gerbing 18.10.2017
                    rc = file_delete(gstrFotosMdbLocation & "\Newfotos.mdb", , True)
                End If
                'MsgBox "Komprimieren der Datenbank wurde ausgeführt"
            Else
                'MsgBox "Komprimieren der Datenbank wurde versucht, aber konnte nicht ausgeführt werden"
            End If
            On Error GoTo 0
            WriteCDE "1"  'CompactDatabaseEnded=1                                                               'Gerbing 08.03.2020
        End If
    End If                                                                                                      'Gerbing 03.05.2020
    Me.MousePointer = vbHourglass                                   'Gerbing 07.11.2011
    Screen.MousePointer = vbHourglass                                   'Gerbing 07.11.2011
        
'    Me.Top = 0                                                      'Gerbing 20.11.2008
'    Me.Left = 0
'    screenWidth = GetDeviceCaps(Me.hDC, HORZRES)                    'Gerbing 29.03.2012
'    screenHeight = GetDeviceCaps(Me.hDC, VERTRES)                   'Gerbing 29.03.2012
'    SGVH = screenWidth / screenHeight                               'Gerbing 24.09.2009
'    'Me.Width = ScreenWidth / 2 * Screen.TwipsPerPixelX             'Gerbing 09.05.2012
'    Me.height = screenHeight * Screen.TwipsPerPixelY

    
    Me.MousePointer = vbNormal                                      'Gerbing 07.11.2011
    '-----------------------------------------------------------------
    If gblnSQLServerVersion = True Then
        gblnVollversion = True
        gblnProversion = True
    End If
    '-----------------------------------------------------------------
    '29.04.2019 und 18.05.2018
    Copyright = Right(FileInfo(AppPath & "\fotos.exe").LegalCopyright, 2) 'Gerbing 18.05.2018 und 29.04.2019
    If Copyright = "-1" Then
        gblnVollversion = True
        gblnProversion = True
    End If
    '----------------------------------------------------------------------------------
'    If gblnErsterStart = True Then                                  'Gerbing 21.11.2019
'        PublicLanguage = "9"                                        'Gerbing 21.11.2019
'        gblnErsterStart = False                                     'Gerbing 21.11.2019
'    End If
    If gblnSQLServerVersion = True Then                             'Gerbing 29.12.2011
        gstrFotosMdbLocation = PublicLocationFotos
    End If
    gstrCommandLine = "/WRITE"                              'Gerbing 11.03.2010
    '-------------------------------------------------------------
    NL = vbNewLine
    Me.MousePointer = vbNormal                                                              'Gerbing 29.07.2007

    'On Error Resume Next               'Gerbing 27.09.2010
    On Error GoTo 0                     'Gerbing 27.09.2010
    '--------------------
    'mnuxxx.Caption mit LoadResString laden                                                 'Gerbing 08.01.2018
    mnuDatei.Caption = LoadResString(3167 + Sprache)
    mnuDateiKopieren.Caption = LoadResString(3168 + Sprache)
    mnuEmailAnhang.Caption = LoadResString(3169 + Sprache)
    mnuExplorer.Caption = LoadResString(3170 + Sprache)
    mnuFeldAktualisierung.Caption = LoadResString(3171 + Sprache)
    mnuFotosmdbStarten.Caption = LoadResString(3172 + Sprache)
    mnuGeoPosition.Caption = LoadResString(3173 + Sprache)
    mnuEinfügenGeoPosition.Caption = LoadResString(3195 + Sprache)                          'Einfügen Geo-Position Gerbing 02.10.2019
    mnuHilfe.Caption = LoadResString(3174 + Sprache)
    mnuHyperlink.Caption = LoadResString(3175 + Sprache)
    mnuImport.Caption = LoadResString(3176 + Sprache)
    mnuKopiereMdb.Caption = LoadResString(3177 + Sprache)                                   'Export der aktuell ausgewählten Dateien...
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
    '----------------------------------------------------------------------------------------------------------
    'Seit Version 15.0.5 muss es in der Tabelle Fotos die Felder GPSLatitude und GPSLongitude geben        'Gerbing 23.10.2019
    'die erzeugt das Programm selbst, wenn es sie noch nicht gibt
    If gblnSchreibgeschützt = False Then
            On Error Resume Next
            SQL = "select GPSLatitude From Fotos;"
            Set rstNeu = New ADODB.Recordset
            With rstNeu
                .Source = SQL
                .ActiveConnection = DBado                                               'Gerbing 23.11.2017
                .CursorType = adOpenDynamic
                .LockType = adLockOptimistic
                .CursorLocation = adUseClient
                .Open
            End With
            If Err.Number <> 0 Then
                'hier existiert das Feld GPSLatitude und GPSLongitude nicht
                If gblnSchreibgeschützt = False Then
                    If gblnSQLServerVersion = True Then
                        'SQL Server                                                     'Gerbing 23.11.2017
                        DBado.Execute _
                            "ALTER TABLE Fotos ADD GPSLatitude FLOAT"          'es heißt ADD und nicht ADD COLUMN
                        DBado.Execute _
                            "ALTER TABLE Fotos ADD GPSLongitude FLOAT"
                    Else
                        'Access Version
                        'also wird Feld GPSLatitude und GPSLongitude erzeugt
                        DBado.Execute _
                            "ALTER TABLE Fotos ADD COLUMN GPSLatitude DOUBLE"
                        DBado.Execute _
                            "ALTER TABLE Fotos ADD COLUMN GPSLongitude DOUBLE"
                    End If
                End If
                rstNeu.Close
                On Error GoTo 0
            End If
    End If
    '----------------------------------------------------------------------------------------------------------
    'Seit Version 15.0.5 muss es in der Tabelle Fotos das Feld EXIFDateTimeOriginal geben        'Gerbing 14.11.2019
    'das erzeugt das Programm selbst, wenn es das noch nicht gibt
    If gblnSchreibgeschützt = False Then
            On Error Resume Next
            SQL = "select EXIFDateTimeOriginal From Fotos;"
            Set rstNeu = New ADODB.Recordset
            With rstNeu
                .Source = SQL
                .ActiveConnection = DBado                                               'Gerbing 23.11.2017
                .CursorType = adOpenDynamic
                .LockType = adLockOptimistic
                .CursorLocation = adUseClient
                .Open
            End With
            If Err.Number <> 0 Then
                'hier existiert das Feld EXIFDateTimeOriginal nicht
                If gblnSchreibgeschützt = False Then
                    If gblnSQLServerVersion = True Then
                        'SQL Server                                                     'Gerbing 23.11.2017
                        DBado.Execute _
                            "ALTER TABLE Fotos ADD EXIFDateTimeOriginal varchar(255)"          'es heißt ADD und nicht ADD COLUMN
                    Else
                        'Access Version
                        'also wird Feld EXIFDateTimeOriginal erzeugt
                        DBado.Execute _
                            "ALTER TABLE Fotos ADD COLUMN EXIFDateTimeOriginal TEXT"
                    End If
                End If
                rstNeu.Close
                On Error GoTo 0
            End If
    End If
    '----------------------------------------------------------------------------------------------------------
    'Seit Version 15.0.5 muss es in der Tabelle Fotos das Feld VideoDuration geben        'Gerbing 14.11.2019
    'das erzeugt das Programm selbst, wenn es das noch nicht gibt
    If gblnSchreibgeschützt = False Then
            On Error Resume Next
            SQL = "select VideoDuration From Fotos;"
            Set rstNeu = New ADODB.Recordset
            With rstNeu
                .Source = SQL
                .ActiveConnection = DBado                                               'Gerbing 23.11.2017
                .CursorType = adOpenDynamic
                .LockType = adLockOptimistic
                .CursorLocation = adUseClient
                .Open
            End With
            If Err.Number <> 0 Then
                'hier existiert das Feld VideoDuration nicht
                If gblnSchreibgeschützt = False Then
                    If gblnSQLServerVersion = True Then
                        'SQL Server                                                     'Gerbing 23.11.2017
                        DBado.Execute _
                            "ALTER TABLE Fotos ADD VideoDuration INT"          'es heißt ADD und nicht ADD COLUMN
                    Else
                        'Access Version
                        'also wird Feld EXIFDateTimeOriginal erzeugt
                        DBado.Execute _
                            "ALTER TABLE Fotos ADD COLUMN VideoDuration LONG"
                    End If
                End If
                rstNeu.Close
                On Error GoTo 0
            End If
    End If
    '----------------------------------------------------------------------------------------------------------
    'lblVideoWirdgeladen.Caption = "Video wird geladen. Drücken Sie F5 zur Anzeige des Listenfensters"                                     'Gerbing 08.11.2005
    lblVideoWirdgeladen.Caption = LoadResString(1013 + Sprache)
    'lblLeereForm.Caption = "Wählen Sie ein anderes Bild, Tasten F2/F3 oder Alt+Pfeil-Tasten"   'Gerbing 08.11.2005
    lblLeereForm.Caption = LoadResString(1012 + Sprache)
    'FotoAlbumTitle = "FotoAlbum-"
    FotoAlbumTitle = LoadResString(1001 + Sprache)
    FotoAlbumTitle = FotoAlbumTitle & gstrFotosMdbLocation          'Gerbing 07.11.2011
    '-----------------------------------------------------------------
    'Jetzt dauert es 20 Sekunden bis query.Form_Load                'Gerbing 27.08.2017
    Query.Hide                              'Jetzt verschwindet Query.Caption auf der Taskleiste
    Form1.Hide                                                      'Gerbing 30.09.2013
    'bis hierher dauert es 40 Sekunden                              'Gerbing 27.08.2017
    Unload frmFotoAlbumWirdGeladen                                  'Gerbing 27.08.2017
    Query.Show          'ohne 1  'Gerbing 15.11.2012                  'Suchkriterien festlegen
    Form1.Hide                                                      'Gerbing 30.09.2013
    gblnErsterStart = False                                         'Gerbing 29.10.2013
    '------------------------------------------------------------------------------------------
    Exit Sub
Fehler:
    Msg = "Errorcode: " & Err & NL
    Msg = Msg & "Errortext: " & Error(Err)
    MsgBox Msg
    End
    Resume
End Sub

Private Sub GeheZumAnfang()
    On Error Resume Next
    frmGridAndThumb.Adodc1.Recordset.MoveFirst
    If Err <> 0 Then
        Msg = "Es wurde kein einziger Datensatz gefunden." & NL
        Msg = Msg & "Mit der F8-Taste können Sie die Suche wiederholen"
        'MsgBox msg                 'Gerbing 08.11.2005
        MsgBox LoadResString(2007 + Sprache) & NL & LoadResString(2008 + Sprache)
    End If
End Sub

Public Sub GeheEinBildZurück()
    Call EinBRückADO
End Sub

Public Sub GeheEinBildVorwärts()
    EinBVorADO
End Sub

Private Sub GeheZumEnde()
    On Error Resume Next
    frmGridAndThumb.Adodc1.Recordset.MoveLast
    If Err = 91 Then    'Objektvariable oder With-Blockvariable nicht festgelegt
        Msg = "Es wurde kein einziger Datensatz gefunden." & NL
        Msg = Msg & "Mit der F8-Taste können Sie die Suche wiederholen"
        'MsgBox msg                 'Gerbing 08.11.2005
        MsgBox LoadResString(2007 + Sprache) & NL & LoadResString(2008 + Sprache)
        Exit Sub
    End If
End Sub

Public Sub BildAnzeigen()
    Dim blnZoomGleich100 As Boolean
    Dim MyGVH As Double               'My-Bild-Größen-Breiten-Verhältnis
    Dim ZoomWidth As Double
    Dim ZoomHeight As Double
    Dim rc As Long

    Dim Msg As String
    Dim DateinamenErweiterung As String
    Dim GesamtPixel As Long
    Dim X As Long
    Dim Y As Long
    Dim n As Long
    Dim pos As Long
    Dim KeyCode As Integer
    Dim Shift As Integer
    Dim temp As String
    Dim temp1 As String
    Dim start As Long
    Dim Dateiname As String
    Dim intLänge As Integer
    
    If gblnComefromVideo = True Then                                'Gerbing 05.08.2013
        Call MediaPlayerStop                                        'Gerbing 19.04.2014
        Form1.Picture1.Picture = LoadPicture("")
        'Form1.Show                                                 '29.03.2015 auskommentiert weil es sonst flackert von einen Video zum nächsten
    Else
        Unload frmVideo                                             'Gerbing 15.11.2012
    End If
    gblnComefromVideo = False                                       'Gerbing 04.12.2012
    Unload KommentarForm                                            'Gerbing 23.10.2012    'Gerbing 28.10.2012
    'frmVideo.WMP.URL = ""  'Gerbing 26.10.2011 kann entfallen, wenn sowieso unload frmvideo folgt
    gblnComeFromBildanzeigen = True                                 'Gerbing 31.12.2012
    Unload Hilfebx                                                  'Gerbing 04.12.2012
        
    On Error Resume Next
    If gblnComeFromThumbs = False Then
        Call FRODateiname
    End If
    If Mid(gstrFRODN, Len(gstrFRODN) - 3, 1) = "." Then             'Gerbing 25.06.2006
        intLänge = 3
        GoTo LängeGefunden                                          'Gerbing 04.10.2012
    End If
    If Mid(gstrFRODN, Len(gstrFRODN) - 4, 1) = "." Then
        intLänge = 4
        GoTo LängeGefunden                                          'Gerbing 04.10.2012
    End If
    If Mid(gstrFRODN, Len(gstrFRODN) - 5, 1) = "." Then
        intLänge = 5
    End If
LängeGefunden:
    DateinamenErweiterung = Right(gstrFRODN, intLänge)
    If Err = 91 Then    'Objektvariable oder With-Blockvariable nicht festgelegt
        Msg = "Es wurde kein einziger Datensatz gefunden." & NL
        Msg = Msg & "Mit der F8-Taste können Sie die Suche wiederholen"
        'MsgBox msg                 'Gerbing 08.11.2005
        MsgBox LoadResString(2007 + Sprache) & NL & LoadResString(2008 + Sprache)
        Exit Sub
    End If
    DateinamenErweiterung = UCase(DateinamenErweiterung)
    
    On Error Resume Next
    Err = 0
    If gblnMouseSichtbar = True Then
        Me.MousePointer = vbHourglass                                                       'Gerbing 29.07.2007
    End If
    Select Case DateinamenErweiterung
        Case "AVI", "MPG", "PEG", "MOV", "MPE", "ASF", "ASX", "WMV", "MP4", "MKV", "FLV"      'Gerbing 10.12.2017
            If StrComp(Right(gstrFRODN, 4), "JPEG", vbTextCompare) = 0 Then     'Gerbing 21.08.2006
                'bei JPEG und zB pdf nichts tun
            Else
                blnTimer1Enabled = False
                DisableTimer                                                                    'Gerbing 29.03.2012
                Call VideoAbspielen
                Exit Sub
            End If
    End Select
    '======================================================================================
    'hierher kommt es nur bei Fotos und allen Nicht-Videos
    Select Case DateinamenErweiterung                                   'Gerbing 11.12.2005
        Case "BMP", "DIB", "EMF", "GIF", "ICO", "JPG", "PNG", "TIF", "TIFF", "WMF"
            'das sind die native mode formats
        Case Else
            'das sind die link mode formats
            Call mnuVerknüpfteAnwendung_Click                           'Gerbing 01.09.2017
            'me.MousePointer = vbNormal
            lblLeereForm.Visible = True                                 'Gerbing 29.03.2012
            Form1.Visible = True                                        'Gerbing 06.06.2015
            Exit Sub
    End Select
    If F6Continous = True Then
        blnTimer1Enabled = True
        EnableTimer lngTimer1Interval
    End If
    On Error Resume Next
    If gblnComeFromThumbs = False Then                                                      'Gerbing 29.03.2015
        Call FRODateiname
    End If
    BHVImageVorige = 0
    If Err.Number <> 0 Then
        Me.MousePointer = vbNormal                                                          'Gerbing 29.07.2007
        Call BildFehler
        Exit Sub
    End If
    '---------------------------------------------------------------------------------------------------------
    If gblnMouseSichtbar = True Then
        If gblnRechteckLupeScharf = True Then                                                   'Gerbing 29.07.2007
            Me.MousePointer = vbCustom                                                      'Gerbing 29.07.2007
            'Me.MouseIcon = LoadPicture(AppPath & "\SquareZoom.ico")                        'Gerbing 29.07.2007
            Me.MouseIcon = LoadResPicture(105, 1)                                           'Gerbing 04.03.2013
        Else
            Me.MousePointer = vbDefault                                                     'Gerbing 29.07.2007
        End If
    Else
        Me.MousePointer = vbCustom      '99                                                 'Gerbing 29.07.2007
        'Me.MouseIcon = LoadPicture(AppPath & "\MOUSE01.ICO")                               'Gerbing 29.07.2007
        Me.MouseIcon = LoadResPicture(104, 1)                                               'Gerbing 04.03.2013
    End If
    '--------------------------------------------------------------------------------------------
    blnComeFromError = False                                                                'Gerbing 15.02.2014
    rc = LoadPicBox(gstrFRODN)
    If rc <> 0 Then
        blnComeFromError = True                                                             'Gerbing 15.02.2014
        MakeGradient Me, vbBlue, vbGreen, GRADIENT_FILL_RECT_V
        MakeGradient Form1.Picture1, vbBlue, vbGreen, GRADIENT_FILL_RECT_V
        'MsgBox gstrFRODN & " Bild kann nicht geladen werden"
        Msg = gstrFRODN & " " & LoadResString(2056 + Sprache)
        MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
        lblLeereForm.Visible = True
        Exit Sub
    End If
    
    'Untersuchungen, ob das Bild größer ist als die Bildschirmauflösung und dessen Konsequenzen
    blnZoomGleich100 = False
    If PublicZoomToFullscreen <> "1" Then                                                       'Gerbing 29.04.2013
        'was schmaler oder niedriger ist als 500 Pixel wird beim ersten Laden nicht vergrößert oder verkleinert
        'glngZoomProzent = 100 heißt anzeigen in Originalgröße
        If gsngPicWidth < 500 Or gsngPicHeight < 500 Then
            blnZoomGleich100 = True
        End If
        If gsngPicWidth <= screenWidth And gsngPicHeight <= screenHeight And PublicZoomToFullscreen = "0" Then
            blnZoomGleich100 = True
        End If
        If gsngPicWidth < 640 And PublicZoomToFullscreen = "640" Then
            blnZoomGleich100 = True
        End If
        If gsngPicWidth < 800 And PublicZoomToFullscreen = "800" Then
            blnZoomGleich100 = True
        End If
        If gsngPicWidth < 1024 And PublicZoomToFullscreen = "1024" Then
            blnZoomGleich100 = True
        End If
        If gsngPicWidth < 1024 And PublicZoomToFullscreen = "1024768" Then
            blnZoomGleich100 = True
        End If
        If gsngPicHeight < 768 And PublicZoomToFullscreen = "1024768" Then                            'Gerbing 29.04.2013
            blnZoomGleich100 = True
        End If
        'was höher oder breiter ist als Screenwidth bzw Screenheight wird beim ersten Laden verkleinert
        If gsngPicWidth > screenWidth Or gsngPicHeight > screenHeight Then
            blnZoomGleich100 = False
        End If
    End If
    
    If blnZoomGleich100 = False Then
        MyGVH = gsngPicWidth / gsngPicHeight
        If MyGVH >= SGVH Then
            ZoomWidth = screenWidth / gsngPicWidth
            glngZoomProzent = 100 * ZoomWidth
        Else
            If Query.chkFensterGrößeÄnderbar.Value = 0 Then
                'Anzeige ohne Titelleiste und Taskbar
                ZoomHeight = screenHeight / gsngPicHeight
            Else
                'Anzeige mit Titelleiste und Taskbar
                'Da muss ich 150 Pixel für die Task Bar im Windows 7 abziehen
                'ZoomHeight = (ScreenHeight - 150) / gsngPicHeight                  'Gerbing 09.05.2012
                ZoomHeight = screenHeight / gsngPicHeight
            End If
            glngZoomProzent = 100 * ZoomHeight
        End If
    Else
        glngZoomProzent = 100
    End If
    glngDiffX = 0
    glngDiffY = 0
    'gblnRechteckLupeScharf = False                                                         'Gerbing 14.10.2014
    
    XYPos.lblBildgröße = gsngPicWidth & " x " & gsngPicHeight
    If gblnMouseSichtbar = True Then
        If gblnRechteckLupeScharf = True Then
            Me.MousePointer = vbCustom
            'Me.MouseIcon = LoadPicture(AppPath & "\SquareZoom.ico")
            Me.MouseIcon = LoadResPicture(105, 1)                                           'Gerbing 04.03.2013
        Else
            Me.MousePointer = vbDefault
        End If
    Else
        Me.MousePointer = vbCustom      '99
        'Me.MouseIcon = LoadPicture(AppPath & "\MOUSE01.ICO")
        Me.MouseIcon = LoadResPicture(104, 1)                                               'Gerbing 04.03.2013
    End If
    Call AbrufenBildPosList
    '-------------------------------------------------------------------------
    If Query.CheckUseAudioComments.Value = 1 Then                                           'Gerbing 26.10.2011
        If gblnVollversion = True Then                                                      'Gerbing 09.12.2009
            'Wenn es eine gleichnamige Audio-Datei zum aktuellen Dateiname gibt, wird
            'sie mit frmStartSoundAutomatisch.MediaPlayer1 abgespielt
            'Call FRODateiname                                                              'Gerbing 25.03.2018
            'Der Dateiname wird ermittelt durch Suchen ab rechtem Rand bis zum Punkt
            start = LenB(gstrFRODN) - 2
            Do
                pos = InStrB(start, gstrFRODN, ".")
                If pos <> 0 Then
                    Dateiname = MidB(gstrFRODN, 1, pos - 1)
                    Exit Do
                End If
                start = start - 1
            Loop
            'strtemp = Dir(AppPath & "\Fotos.mdb")
            'If file_path_exist(AppPath & "\Fotos.mdb") = False Then
            If file_path_exist(Dateiname & ".mp3") = True Then
                frmStartSoundAutomatisch.MediaPlayer1.url = Dateiname & ".mp3"
                frmStartSoundAutomatisch.MediaPlayer1.Controls.play                        'Gerbing 09.12.2009
            End If
            If file_path_exist(Dateiname & ".wav") = True Then
                frmStartSoundAutomatisch.MediaPlayer1.url = Dateiname & ".wav"                 'Gerbing 01.09.2008
                frmStartSoundAutomatisch.MediaPlayer1.Controls.play                        'Gerbing 09.12.2009
            End If
        End If
    End If
    '-------------------------------------------------------------------------
    If Query.chkFensterGrößeÄnderbar.Value = 1 Then                         'Gerbing 04.09.2012
        Form1.WindowState = 2                                              'Gerbing 29.04.2013 27.09.2017
'        Me.Width = Form1.Width
'        Me.Height = Form1.Height
        On Error Resume Next                                                'Gerbing 10.06.2012
    'Form1.Show                                                                             'Gerbing 29.04.2013
        Picture1.width = Form1.width
        Picture1.height = Form1.height
    Else
        On Error Resume Next
        Form1.WindowState = 2       '0=zuerst normal, sonst kommt Fehler bei Form1.Width einstellen 'Gerbing 17.10.2014=2
        Form1.width = screenWidth * Screen.TwipsPerPixelX
        Form1.height = screenHeight * Screen.TwipsPerPixelY
        Picture1.width = Form1.width
        Picture1.height = Form1.height
        Form1.WindowState = 2       '2=maximiert wenn die Taskleiste automatisch ausgeblendet wird, soll das Bild bis ganz unten hin gehen
        'Form1.Show                                                                             'Gerbing 29.04.2013
    End If
    On Error GoTo 0
    Call Form1.MyDrawImage(gstrFRODN, glngZoomProzent)                      'Gerbing 10.06.2012
    'Nach dem Anzeigen des Bildes muss abgefragt werden, ob ein Kommentar eingeblendet werden soll
    Call KommentarNachBildanzeigen                                          'Gerbing 04.05.2015
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Call Hilfebox
    End If
End Sub

Public Sub Hilfebox()
    If Query.chkFensterGrößeÄnderbar.Value = 1 Then            'Gerbing 06.12.2005
        Hilfebx.Top = Form1.Top                         'Gerbing 06.12.2006
        Hilfebx.Left = Form1.Left
    End If
    Hilfebx.Show
    On Error Resume Next
    'AppActivate FotoAlbumTitle
End Sub

Public Sub Form_Resize()
    Dim OsVersInfo As OSVERSIONINFO                     'Gerbing 19.11.2012
    
    If Form1.WindowState = vbMinimized Then
        'MsgBox "Form1.Form_Resize " & "Windowstate=" & Form1.WindowState                   'Gerbing 04.05.2015
        frmGridAndThumb.WindowState = 1
        XYPos.WindowState = 1
    End If
    If gblnComefromVideo = True Then Exit Sub           'Gerbing 19.09.2012
    If blnComeFromError = True Then                     'Gerbing 15.02.2014
        MakeGradient Me, vbBlue, vbGreen, GRADIENT_FILL_RECT_V
        MakeGradient Form1.Picture1, vbBlue, vbGreen, GRADIENT_FILL_RECT_V
        'MsgBox gstrFRODN & " Bild kann nicht geladen werden"
        Msg = gstrFRODN & " " & LoadResString(2056 + Sprache)
        MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
        Exit Sub                                        'Gerbing 15.02.2014
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Query.Beenden
'    End

'    If lngPointer Then
'        retcode = GdipDisposeImage(lngPointer)
'        lngPointer = 0                                                                  'Gerbing 19.04.2017
'    End If
'    If m_lngGraphics Then
'        If GdipDeleteGraphics(m_lngGraphics) Then
'            'MsgBox "Graphics object could not be deleted", vbCritical                  'Gerbing 19.11.2013
'        End If
'    End If
'    GdiplusShutdown m_lngInstance
    
    'Unload Hilfebx
    Set EXF = Nothing                                                                   'Gerbing 07.05.2007
    Set IniFso = Nothing
    'Unload frmGridAndThumb                                                             'Gerbing 29.03.2015 auskommentiert
    Unload Hilfebx
    Unload KommentarForm
    Unload Query
    'Unload QueryJedesFeld
    Unload MP
    End


'    If lngPointer Then
'        retcode = GdipDisposeImage(lngPointer)
'    End If
'    If m_lngGraphics Then
'        If GdipDeleteGraphics(m_lngGraphics) Then
'            'MsgBox "Graphics object could not be deleted", vbCritical                  'Gerbing 19.11.2013
'        End If
'    End If
'    GdiplusShutdown m_lngInstance
'    'Unload Hilfebx
'    Set EXF = Nothing                                                                   'Gerbing 07.05.2007
'    Set IniFso = Nothing
'    'Unload frmGridAndThumb                                                             'Gerbing 29.03.2015 auskommentiert
'    Unload Hilfebx
'    Unload KommentarForm
'    Unload Query
'    'Unload QueryJedesFeld
'    Unload MP
'    End
End Sub

Private Sub lblLeereForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Call Hilfebox
    End If
End Sub

Private Sub lblVideoWirdgeladen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Call Hilfebox
    End If
End Sub

Private Sub mnuDateiKopieren_Click()                                                        'Gerbing 01.09.2017
    frmZwischenablageOderOrdner.Show 1                                                      'Gerbing 11.08.2017
End Sub

Private Sub mnuEinfügenGeoPosition_Click()
    'das Einfügen der Geo-Position geht nicht nur für JPG files sondern für alle
    'aber nur wenn es die Felder GPSLatitude und GPSLongitude gibt
    Dim rc As Integer
    
    rc = GPSFelderPrüfen                                                                        'Gerbing 02.10.2019
    If rc = 0 Then Exit Sub                                                                     'Gerbing 02.10.2019
    frmGPSInDatenbankEintragen.Show 1                                                           'Gerbing 02.10.2019
    frmGridAndThumb.btnRefresh_Click                                                            'Gerbing 20.10.2019
End Sub

Private Sub mnuEmailAnhang_Click()
    Call EmailMitAnhangSenden                                                               'Gerbing 01.09.2017
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
    'das geht nur für JPG files, in anderen ist keine GEO-Positionen vorhanden
    'und nur wenn es die Felder GPSLatitude und GPSLongitude gibt
    Dim rc As Integer
    
    rc = GPSFelderPrüfen                                                                        'Gerbing 02.10.2019
    If rc = 0 Then Exit Sub                                                                     'Gerbing 02.10.2019
    frmFeldAktualisierung.Show 1
    frmGridAndThumb.btnRefresh_Click                                                            'Gerbing 20.10.2019
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

Private Sub mnuGeoPosition_Click()                                                              'Gerbing 01.09.2017
    If Not (gblnVollversion = True And gblnProversion = True) Then              'Gerbing 27.09.2016
        Msg = LoadResString(2335 + Sprache) 'Für diese Funktion benötigen Sie die Professional Version.
        MsgBox Msg
        Exit Sub
    End If
    Call ZeigeGEOPosition                                                                       'Gerbing 03.10.2016
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

Private Sub mnuImport_Click()
    If gblnSQLServerVersion = True Then
        'MsgBox "Diese Funktion wird beim SQL-Server nicht unterstützt"
        MsgBox LoadResString(1829 + Sprache)
    Else
        Call MediaPlayerStop
        ImportForm.Show 1
    End If
End Sub

Private Sub mnuKopiereMdb_Click()                                               'Gerbing 01.09.2017
    Call MediaPlayerStop
    'me.MousePointer = vbDefault                                                'Gerbing 29.07.2007
    blnUnloadExportForm = False
    ExportForm.Show 1
    frmVideo.lblLeereForm.Visible = True                                        'Gerbing 04.12.2012
End Sub

Private Sub mnuLöschen_Click()                                                  'Gerbing 01.09.2017
    Dim antwort As Long
    Dim KeyCode As Integer
    Dim Shift As Integer
    
    SQL = " SELECT *"
'    SQL = SQL & " FROM " & "Fotos"
'    SQL = SQL & " WHERE Merker<>0
'    SQL = SQL & " ORDER BY Dateiname" & ";"
    SQL = " SELECT *"
    SQL = SQL & " FROM Fotos"
    SQL = SQL & " WHERE " & LoadResString(2524 + Sprache) & "<>0"
    SQL = SQL & " ORDER BY " & LoadResString(1028 + Sprache) & ";"
    On Error Resume Next                                                        'Gerbing 12.02.2018
    rstsql.Close                                                                'Gerbing 12.02.2018
    On Error GoTo 0                                                             'Gerbing 12.02.2018
    With rstsql
        .Source = SQL
        .ActiveConnection = DBado                                               'Gerbing 23.11.2017
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
    
'    'Wenn ich mit dem aktuell angezeigten Bild nach Renammdb verzweige, kann ich das aktuell angezeigte Bild weder löschen noch Namen ändern
'    'Ich muss das aktuell angezeigte Bild erst entladen                         'Gerbing 25.03.2018
'    Picture1.Picture = LoadPicture("")
'    If lngPointer Then
'        retcode = GdipDisposeImage(lngPointer)
'        lngPointer = 0
'    End If
'    If m_lngGraphics Then
'        If GdipDeleteGraphics(m_lngGraphics) Then
'            'MsgBox "Graphics object could not be deleted", vbCritical
'        End If
'    End If
'    GdiplusShutdown m_lngInstance                                               'Gerbing 25.03.2018
    
    Do Until rstsql.EOF
        Call LöschenInDatenbankUndStandort(rstsql.Fields(LoadResString(1028 + Sprache)), rstsql) '1028=Dateiname
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
    Dim zähler As Long
    
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
            .ActiveConnection = DBado                                               'Gerbing 23.11.2017
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
'        Msg = "Anzahl zu ändernde Dateinamen=" & rstsql.RecordCount & vbNewLine
'        Msg = Msg & "Wollen Sie diese wirklich ändern?" & vbNewLine
        Msg = LoadResString(2353 + Sprache) & rstsql.RecordCount & vbNewLine
        Msg = Msg & LoadResString(2354 + Sprache) & vbNewLine
        antwort = MsgBox(Msg, vbYesNo + vbDefaultButton1)
        If antwort = vbNo Then
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
            'antwort = yes
            'Das aktuell gezeigte Foto/Video kann nicht umgenannt werden, deshalb entladen
            Picture1.Picture = LoadPicture("")
            If lngPointer Then                                                  'Gerbing 31.12.2012
                retcode = GdipDisposeImage(lngPointer)
                lngPointer = 0                                                  'Gerbing 19.04.2017
            End If
            If m_lngGraphics Then                                               'Gerbing 31.12.2012
                If GdipDeleteGraphics(m_lngGraphics) Then
                    'MsgBox "Graphics object could not be deleted", vbCritical
                End If
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
                rc = NameAs(DateinameAlt, strPath & DateinameNeu & "." & DateinamenErweiterung)
                'jetzt untersuchen, ob rename altername As neuername erfolgreich war
                If rc <> 0 Then
                    If rc = 72 Then 'Datei nicht erzeugt werden, wenn es diese bereits gibt
                        Msg = LoadResString(2360 + Sprache) & zähler & vbNewLine    '2360=Erfolgreich umgenannte Dateien=
                        Msg = Msg & LoadResString(2355 + Sprache) & vbNewLine       '2355=Rename war nicht erfolgreich
                        Msg = Msg & LoadResString(2356 + Sprache) & DateinameAlt & vbNewLine    '2356=altname=
                        Msg = Msg & LoadResString(2357 + Sprache) & strPath & DateinameNeu & "." & DateinamenErweiterung & vbNewLine    '2357=neuname=
                        Msg = Msg & LoadResString(2361 + Sprache) & vbNewLine       '2361=Datei kann nicht erzeugt werden, wenn es diese bereits gibt
                        Msg = Msg & LoadResString(2358 + Sprache) & vbNewLine       '2358=Machen Sie die Änderung mit RenamMdb.exe
                        Msg = Msg & LoadResString(2139 + Sprache)                   '2139=das Programm wird beendet
                        MsgBox Msg
                        End
                    Else
                        'rc=92 oder rc= 98
                        'die aktuelle Datei kann nicht umgenannt werden
                        Msg = LoadResString(2355 + Sprache) & vbNewLine             '2355=Rename war nicht erfolgreich
                        Msg = Msg & LoadResString(2356 + Sprache) & DateinameAlt & vbNewLine    '2356=altname=
                        Msg = Msg & LoadResString(2357 + Sprache) & strPath & DateinameNeu & "." & DateinamenErweiterung & vbNewLine    '2357=neuname=
                        Msg = Msg & LoadResString(2359 + Sprache) & vbNewLine       '2359=Diese Datei kann nicht umgenannt werden
                        Msg = Msg & LoadResString(2358 + Sprache) & vbNewLine       '2358=Machen Sie die Änderung mit RenamMdb.exe
                        MsgBox Msg
                    End If
                Else
                    'rc=0
                    rstsql.Fields(LoadResString(1031 + Sprache)) = DateinameNeu & "." & DateinamenErweiterung   '1031=DateinameKurz
                    DateinameNeuMitPlus = Replace(strPath & DateinameNeu & "." & DateinamenErweiterung, gstrFotosMdbLocation & "\", "+:\")
                    rstsql.Fields(LoadResString(1028 + Sprache)) = DateinameNeuMitPlus                          '1028=Dateiname
                    ThumbnameAlt = strPath & "GerbingThumbs\" & strFile & "." & DateinamenErweiterung & ".jpg"
                    ThumbnameNeu = strPath & "GerbingThumbs\" & DateinameNeu & "." & DateinamenErweiterung & ".jpg"
                    rc = NameAs(ThumbnameAlt, ThumbnameNeu)
                    zähler = zähler + 1
                    'Wenn das foto/video erfolgreich umgenannt werden konnte, muss jetzt gefragt werden, ob es
                    'gleichnamige .wav oder .mp3 files gibt
                    Call AudioDateiMitUmnennen(DateinameAlt, strPath & DateinameNeu & "." & DateinamenErweiterung)  'Gerbing 26.03.2018
                End If
                rstsql.Update
rstsqlMoveNext:
                rstsql.MoveNext
            Loop
            rstsql.Close
        End If
    Else
        'im schreibgeschützten Zustand nicht möglich
        Msg = gstrFotosMdbLocation & "\Fotos.mdb" & vbNewLine
        Msg = Msg & LoadResString(2210 + Sprache)       '2210=Die Datenbank ist schreibgeschützt, Änderungen sind nicht möglich
        MessageBoxW 0, StrPtr(Msg), StrPtr(LoadResString(1119 + Sprache)), vbInformation    '1119=GERBING Fotoalbum 15
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    Msg = LoadResString(2360 + Sprache) & zähler & vbNewLine                        '2360=Erfolgreich umgenannte Dateien=
    Msg = Msg & LoadResString(1007 + Sprache) & " " & LoadResString(2139 + Sprache) '1007=Fertig 2139=Das Programm wird beendet
    MsgBox Msg
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
    
'    'Wenn ich mit dem aktuell angezeigten Bild nach Renammdb verzweige, kann ich das aktuell angezeigte Bild weder löschen noch Namen ändern
'    'Ich muss das aktuell angezeigte Bild erst entladen                         'Gerbing 25.03.2018
'    Picture1.Picture = LoadPicture("")
'    If lngPointer Then
'        retcode = GdipDisposeImage(lngPointer)
'        lngPointer = 0
'    End If
'    If m_lngGraphics Then
'        If GdipDeleteGraphics(m_lngGraphics) Then
'            'MsgBox "Graphics object could not be deleted", vbCritical
'        End If
'    End If
'    GdiplusShutdown m_lngInstance                                               'Gerbing 25.03.2018
    
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
        If gstrRowColChangeName = "" Then                                       'Gerbing 09.10.2014
            gstrRowColChangeName = frmGridAndThumb.Adodc1.Recordset.Fields(LoadResString(1028 + Sprache))
        End If
        cmdline = cmdline & "gstrRowColChangeName=" & gstrRowColChangeName & ";"
        AppId = Shell(AppPath & "\RenamMdb.exe" & " " & cmdline, vbNormalFocus)
        AppActivate AppId
    End If
End Sub

Private Sub mnuRenammdbStarten_Click()
    '3182 = RenamMdb starten(als Tool)
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

    
'    SQL = " SELECT *"
'    SQL = SQL & " FROM " & "Fotos"
'    SQL = SQL & " WHERE Merker<>0
'    SQL = SQL & " ORDER BY Dateiname" & ";"
    SQL = " SELECT *"
    SQL = SQL & " FROM Fotos"
    SQL = SQL & " WHERE " & LoadResString(2524 + Sprache) & "<>0"
    SQL = SQL & " ORDER BY " & LoadResString(1028 + Sprache) & ";"
    If rstsql Is Nothing Then                                                   'Gerbing 23.11.2017
        Set rstsql = New ADODB.Recordset
    Else
        On Error Resume Next
        rstsql.Close
        On Error GoTo 0
    End If
    With rstsql
        .Source = SQL
        .ActiveConnection = DBado                                               'Gerbing 23.11.2017
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
            .ActiveConnection = DBado                                               'Gerbing 23.11.2017
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
            .ActiveConnection = DBado                                               'Gerbing 23.11.2017
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

Private Sub TimerKeyboardHook_Timer()
    Dim X As Long
  
    
    For X = 48 To 90
        If CompKey(X, UCase(Chr$(X))) Then Exit Sub
        If CompKey(X + 48, UCase("NUM " & Chr$(X))) Then Exit Sub
    Next X
    
    If CompKey(8, "BACKSPACE") Then Exit Sub
    If CompKey(9, "TAB") Then Exit Sub
    If CompKey(13, "ENTER") Then Exit Sub
    If CompKey(16, "SHIFT") Then Exit Sub
    If CompKey(17, "STRG") Then Exit Sub
    If CompKey(18, "ALT") Then Exit Sub
    If CompKey(19, "PAUSE") Then Exit Sub
    If CompKey(27, "ESC") Then Exit Sub
    If CompKey(33, "PAGE UP") Then Exit Sub
    If CompKey(34, "PAGE DOWN") Then Exit Sub
    If CompKey(35, "ENDE") Then Exit Sub
    If CompKey(36, "POS1") Then Exit Sub
    If CompKey(37, "LEFT") Then Exit Sub
    If CompKey(38, "UP") Then Exit Sub
    If CompKey(39, "RIGHT") Then Exit Sub
    If CompKey(40, "DOWN") Then Exit Sub
    If CompKey(44, "DRUCK") Then Exit Sub
    If CompKey(45, "INSERT") Then Exit Sub
    If CompKey(46, "DEL") Then Exit Sub
    If CompKey(144, "NUM") Then Exit Sub
    If CompKey(145, "ROLLEN") Then Exit Sub
    
    For X = 112 To 127
        If CompKey(X, "F" & CStr(X - 111)) Then
            Exit Sub
        End If
    Next X
    
    ' usw... usw...
End Sub

Private Sub TimerNachFormLoad_Timer()
    'Nur weil sich während Form_Load keine Operation MoveFirst machen läßt
    Dim DateinamenErweiterung As String

    TimerNachFormLoad.Enabled = False
    'If Query.OKGewählt = False Then Exit Sub                    'Gerbing 06.12.2005
    
    On Error Resume Next
    frmGridAndThumb.Adodc1.Recordset.MoveFirst
    If Err <> 0 Then
        If gblnWeiterMitLeererDatenbank = True Then
            'msg = "Mit den Tasten Strg + I können Sie den Import starten"
            Msg = LoadResString(2225 + Sprache)
            MsgBox Msg
            Exit Sub
        Else
            'Msg = "Es wurde kein einziger Datensatz gefunden." & NL
            'Msg = Msg & "Mit der F8-Taste können Sie die Suche wiederholen"
            Msg = LoadResString(2007 + Sprache) & NL & LoadResString(2008 + Sprache)
            MsgBox Msg
            Exit Sub
        End If
    End If
    On Error GoTo 0                                             'Gerbing 31.12.2005
'    frmGridAndThumb.Adodc1.Recordset.MoveLast                         'Gerbing 16.06.2005 'Gerbing 25.06.2013
'    frmGridAndThumb.Adodc1.Recordset.MoveFirst
    Query.RecordCount = frmGridAndThumb.Adodc1.Recordset.RecordCount
    Call BildAnzeigen
    If gblnComefromVideo = True Then                                                            'Gerbing 02.04.2018
        If Query.chkFensterGrößeÄnderbar.Value = 1 Then                                         'Gerbing 21.05.2012
            'Achtung in der IDE wird unicode in Form.Caption nicht angezeigt
            ShowTitleBar True, False                                     'Gerbing 04.09.2012 'taskbar visible, Video
        Else
            'Achtung in der IDE wird unicode in Form.Caption nicht angezeigt
            ShowTitleBar False, False                                    'Gerbing 04.09.2012 'taskbar unvisible, Video
        End If
    Else
        If Query.chkFensterGrößeÄnderbar.Value = 1 Then                                         'Gerbing 21.05.2012
            'Achtung in der IDE wird unicode in Form.Caption nicht angezeigt
            ShowTitleBar True, True                                     'Gerbing 04.09.2012 'taskbar visible, Foto
        Else
            'Achtung in der IDE wird unicode in Form.Caption nicht angezeigt
            ShowTitleBar False, True                                    'Gerbing 04.09.2012 'taskbar unvisible, Foto
        End If
    End If
    If gblnComefromVideo = True Then                                'Gerbing 19.09.2012
        frmVideo.Show
    Else
        Form1.Show
    End If
End Sub

Private Sub BildFehler()
    Dim Msg As String
    
    On Error Resume Next
    'msg = "Bild kann nicht geladen werden." & NL
    Msg = LoadResString(2056 + Sprache) & NL
    'msg = msg & "Bezeichnung in der Datenbank=" & frmGridAndThumb.Adodc1.Recordset("Dateiname") & NL
    Msg = Msg & LoadResString(2057 + Sprache) & frmGridAndThumb.Adodc1.Recordset(LoadResString(1028 + Sprache)) & NL
    If gblnSQLServerVersion = True Then
        'msg = msg & "Bezeichnung nach Ersetzen mit " & PublicLocationFotos & "=" & gstrFRODN & NL
        Msg = Msg & LoadResString(1822 + Sprache) & PublicLocationFotos & "=" & gstrFRODN & NL
    Else
        'msg = msg & "Bezeichnung nach Ersetzen mit " & gstrFotosMdbLocation & "=" & gstrFRODN & NL
        Msg = Msg & LoadResString(1822 + Sprache) & gstrFotosMdbLocation & "=" & gstrFRODN & NL
    End If
    'msg = msg & "Prüfen Sie, ob diese Datei existiert." & NL & NL
    Msg = Msg & LoadResString(2059 + Sprache) & NL & NL
    
    If gblnSQLServerVersion = True Then
        'msg = msg & "Prüfen Sie, ob die Angabe " & PublicLocationFotos  & " richtig ist" & vbNewLine & vbNewLine
        Msg = Msg & LoadResString(1820 + Sprache) & PublicLocationFotos & LoadResString(1821 + Sprache) & vbNewLine & vbNewLine
        'MsgBox Msg, vbInformation
        MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
    Else
        'msg = msg & "Wenn dieser Fehler alle Dateien Ihrer Datenbank fotos.mdb betrifft," & NL
        Msg = Msg & LoadResString(2060 + Sprache) & NL
        'msg = msg & "kann die Ursache darin zu suchen sein, dass Sie die 3-Einigkeits-Forderung" & NL
        Msg = Msg & LoadResString(2061 + Sprache) & NL
        'msg = msg & "nicht eingehalten haben." & NL
        Msg = Msg & LoadResString(2062 + Sprache) & NL
        'msg = msg & "Seit Version 12.0.0.0 verlangt das Programm die 3-Einigkeit, d.h. dass alle Fotos oder Videos oder andere Dateien" & vbNewLine
        Msg = Msg & LoadResString(2063 + Sprache) & vbNewLine
        'msg = msg & "unterhalb von gstrFotosMdbLocation stehen. AppPath ist der Name des Ordners in dem fotos.exe steht." & vbNewLine
        Msg = Msg & LoadResString(2064 + Sprache) & vbNewLine
        'MsgBox Msg, vbInformation
        MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
    End If
    Form1.lblLeereForm.Top = Form1.height / 2 - Form1.lblLeereForm.height / 2
    Form1.lblLeereForm.Left = Form1.width / 2 - Form1.lblLeereForm.width / 2
    Form1.lblLeereForm.Visible = True
End Sub


Public Sub VideoAbspielen()
    Dim RetVal As Long
    Dim MM As New MovieModule
    Dim Error As Long
    Dim MyScreenWidth As Long
    Dim MyScreenHeight As Long
    Dim zähler As Long                                                      'Gerbing 25.04.2012
    Dim KeyCode As Integer
    Dim Shift As Integer

    SammelTextGetAsyncKeyState = ""                                         'Gerbing 29.03.2015
    gblnComeFromBildanzeigen = False                                        'Gerbing 26.09.2013
    Picture1.Picture = LoadPicture("")                                      'Sonst ist beim Video Abspielen zusätzlich das zuletzt gezeigte Bild sichbar
    Form1.lblLeereForm.Visible = False
    frmVideo.lblLeereForm.Visible = False                                   'Gerbing 16.06.2012
    lblLeereForm.Top = Form1.height / 2 - lblLeereForm.height / 2           'Gerbing 11.12.2005
    lblLeereForm.Left = Form1.width / 2 - lblLeereForm.width / 2
    MM.Filename = gstrFRODN                                                 'Gerbing 03.07.2012
    '----------------------------------
    'Abspielen mit externem Mediaplayer
    If PublicPlayVideosWith <> "10" Then                  'Gerbing 29.03.2012
        Form1.Hide                                      'Gerbing 23.10.2013
        frmVideo.Show                                   'Gerbing 23.10.2013
        gblnComefromVideo = True                        'Gerbing 15.11.2012
        frmVideo.WMP.Visible = False                    'Gerbing 15.11.2012
        frmVideo.lblLeereForm.Visible = True            'Gerbing 15.11.2012
        frmVideo.WMP.settings.autoStart = False         'Gerbing 26.10.2011
        frmVideo.WMP.url = gstrFRODN                        'Gerbing 26.10.2011
        If frmVideo.WMP.url = "" Then                   'Gerbing 26.10.2011
            Call VideoFehler
            Exit Sub
        End If
        If gstrWmplayerFolder = "" Then                 'Gerbing 18.01.2013
            Msg = "error - replace the settings - videos" & vbNewLine
            Msg = Msg & "external video player not found"
            MsgBox Msg
            Exit Sub
        End If
        If F6Continous = True Then                              'Gerbing 11.02.2013
            glngVideoDuration = str(MM.getLengthInSec)                              'Gerbing 06.05.2013
'            glngStartMillisek = timeGetTime
            Error& = MM.openMovieWindow(Me.hWnd, "child")                           'Gerbing 03.07.2012
            Error& = MM.extractDefaultMovieSize(wancho, walto)                      'Gerbing 03.07.2012
'            glngEndMillisek = timeGetTime
'            Debug.Print "EndMillisec=" & glngEndMillisek
'            Debug.Print "Millisekunden für extractDefaultMovieSize" & "=" & (glngEndMillisek - glngStartMillisek)
            TimerVideoDuration.Enabled = True
        End If
        frmVideo.WMP.url = ""
        'wmplayer kann durch beispielsweise folgende command line gestartet werden
        'C:\Programme\Windows Media Player\wmplayer.exe Dateiname /fullscreen
        'doppelte Hochkomma sind nötig, wenn der Dateiname Leerzeichen enthält
        gvarVideoAppid = Shell(gstrWmplayerFolder & " " & """" & gstrFRODN & """", vbNormalFocus)    'Gerbing 19.10.2007
        DoEvents
        On Error Resume Next
        'Call ErsatzFürAppActivate("WMP Skin Host", "Windows Media Player")
        AppActivate gvarVideoAppid, False
        'AppActivate "Windows Media Player", False
        On Error GoTo 0
        If F6Continous = True Then
            frmGridAndThumb.Hide
        Else
        Call AnpassenNutzerWunsch(Me)                                                       'Gerbing 11.03.2017
            Call AnpassenHeadFont(frmGridAndThumb.DBGridNeu)                                'Gerbing 23.06.2011
            'frmGridAndThumb.Show                                                           'Gerbing 21.05.2012
        End If
        Exit Sub
    End If
    '----------------------------------------------------------------------
    'Abspielen mit internem Mediaplayer
    frmVideo.Show  ' auskommentiert Gerbing 27.11.2016                                      'Gerbing 10.06.2012
    frmGridAndThumb.Hide   'auskommentiert 27.11.2016                                       'Gerbing 21.07.2005
    If gblnWasOptThumbClick = False Then
        Call FRODateiname
    End If
    'frmVideo.WMP.uiMode = "invisible"      'auskommentiert 23.10.2013                      'Gerbing 01.09.2008
    frmVideo.WMP.settings.autoStart = False                                                 'Gerbing 01.09.2008
    frmVideo.WMP.width = 1
    frmVideo.WMP.url = gstrFRODN                                                            'Gerbing 01.09.2008
    frmVideo.WMP.Visible = True     'erst nach frmVideo.WMP.URL = ...27.11.2016                                                             'Gerbing 01.09.2008
    MyScreenWidth = GetDeviceCaps(Me.hDC, HORZRES) * Screen.TwipsPerPixelX
    MyScreenHeight = GetDeviceCaps(Me.hDC, VERTRES) * Screen.TwipsPerPixelY
    On Error Resume Next
    Err = 0
    frmVideo.WMP.Controls.play
    TimerKeyboardHook.Enabled = True                                                        'Gerbing 29.03.2015

    'MsgBox "Videoabspielen:" & frmVideo.WMP.URL
    If Err <> 0 Then
        On Error GoTo 0
        Call VideoFehler
        Form1.Hide                                                                          'Gerbing 05.08.2013
        Exit Sub
    End If
    On Error GoTo 0
    zähler = 0
    frmVideo.WMP.Left = 0
    If gblnBildBeschreibung = False Then                                                    'Gerbing 030.03.2012
        frmVideo.WMP.Top = 0
        frmVideo.txtBildbeschreibung.Visible = False
    Else
        frmVideo.WMP.Top = txtBildbeschreibung.height
        frmVideo.txtBildbeschreibung.width = Me.width
        frmVideo.txtBildbeschreibung.Visible = True
        frmVideo.txtBildbeschreibung.Text = gstrFRODN
    End If
    'Bild oben anordnen
    '735 Twips ist frmVideo.WMP mindestens hoch auch wenn kein Platz für ein Bild ist
    frmVideo.WMP.uiMode = "full"                                                            'Gerbing 01.09.2008
    
    frmVideo.WMP.stretchToFit = True                                                        'Gerbing 26.11.2012
    frmVideo.WMP.width = Form1.width - 300                          'Gerbing 03.07.2012         10.10.2013
    frmVideo.WMP.height = Form1.height                          'Gerbing 04.02.2013
    If Query.chkFensterGrößeÄnderbar.Value = 1 Then
        frmVideo.WMP.height = frmVideo.WMP.height - 735
    End If
    '----------------------------------
    lblVideoWirdgeladen.Visible = False
    Me.MousePointer = vbDefault                                                             'Gerbing 29.07.2007
    gblnComefromVideo = True                                                                'Gerbing 19.09.2012
    If Query.chkFensterGrößeÄnderbar = 1 Then          '1=aktiviert
        On Error Resume Next
        Me.Left = 0                             'Gerbing 29.03.2012
        Me.Top = 0
        On Error GoTo 0
        ShowTitleBar True, False                          'Gerbing 04.09.2012   'taskbar visible, Video
    Else
        ShowTitleBar False, False                         'Gerbing 04.09.2012   'taskbar unvisible, Video
    End If
    Form1.Hide                                                                              'Gerbing 05.08.2013
    Sleep 100                                                                               'Gerbing 01.09.2008
    If F6Continous = True Then
        TimerToPlayVideo.Enabled = True                                                     'Gerbing 05.08.2013
    End If
End Sub

Private Sub VideoFehler()
    Dim Msg As String
    
'    Msg = Dir(gstrFRODN)
'    If Msg = "" Then
    If file_path_exist(gstrFRODN) = False Then
        'msg = "Videoclip kann nicht geladen werden." & NL
        Msg = LoadResString(2071 + Sprache) & NL
        'msg = msg & "Bezeichnung in der Datenbank=" & frmGridAndThumb.Adodc1.Recordset("Dateiname") & NL
        Msg = Msg & LoadResString(2072 + Sprache) & frmGridAndThumb.Adodc1.Recordset(LoadResString(1028 + Sprache)) & NL
        'msg = msg & "Bezeichnung nach Ersetzen mit gstrFotosMdbLocation=" & gstrFRODN & NL
        Msg = Msg & LoadResString(2058 + Sprache) & gstrFRODN & NL
        'msg = msg & "Prüfen Sie, ob diese Datei existiert." & NL & NL
        Msg = Msg & LoadResString(2059 + Sprache) & NL & NL
        
        'msg = msg & "Wenn dieser Fehler alle Dateien Ihrer Datenbank fotos.mdb betrifft," & NL
        Msg = Msg & LoadResString(2060 + Sprache) & NL
        'msg = msg & "kann die Ursache darin zu suchen sein, dass Sie die 3-Einigkeits-Forderung" & NL
        Msg = Msg & LoadResString(2061 + Sprache) & NL
        'msg = msg & "nicht eingehalten haben." & NL
        Msg = Msg & LoadResString(2062 + Sprache) & NL
        'msg = msg & " 12.0.0.0 verlangt das Programm die 3-Einigkeit, d.h. dass alle Fotos oder Videos oder andere Dateien" & vbNewLine
        Msg = Msg & LoadResString(2063 + Sprache) & vbNewLine
        'msg = msg & "unterhalb von gstrFotosMdbLocation stehen. AppPath ist der Name des Ordners in dem fotos.exe steht." & vbNewLine
        Msg = Msg & LoadResString(2064 + Sprache) & vbNewLine
    Else
        'msg = "Videoclip kann nicht geladen werden." & NL
        Msg = LoadResString(2071 + Sprache) & NL
        Msg = Msg & gstrFRODN & NL
        'msg = msg & "weil vermutlich kein geeigneter Video-Codec auf diesem PC existiert." & NL & NL
        Msg = Msg & LoadResString(2073 + Sprache) & NL & NL
    
        'msg = msg & "Wenn das betroffene Video außerhalb dieses Programm auch nicht abgespielt werden kann," & NL
        Msg = Msg & LoadResString(2074 + Sprache) & NL
        'msg = msg & "müssen Sie sich einen geeigneten Video-Codec besorgen" & NL & NL
        Msg = Msg & LoadResString(2075 + Sprache) & NL & NL
        
        'msg = msg & "Beachten Sie, daß MOV-Videos nur dann abgespielt werden können, wenn sie sehr alt sind." & NL
        Msg = Msg & LoadResString(2076 + Sprache) & NL
        'msg = msg & "Nur Quicktime Versionen bis 2.0 und älter." & NL
        Msg = Msg & LoadResString(2077 + Sprache) & NL
        'msg = msg & "Neuere MOV-Videos konvertieren Sie am besten ins AVI-Format." & NL
        Msg = Msg & LoadResString(2078 + Sprache) & NL
        'msg = msg & "Es gibt dafür die Freeware RAD Video Tools."
        Msg = Msg & LoadResString(2079 + Sprache)
    End If
    'MsgBox Msg, vbInformation
    MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
    lblLeereForm.Top = Form1.height / 2 - lblLeereForm.height / 2                               'Gerbing 11.12.2005
    lblLeereForm.Left = Form1.width / 2 - lblLeereForm.width / 2
    lblLeereForm.Visible = True
    Call MediaPlayerStop
End Sub

Public Sub MediaPlayerStop()
'    Dim rc As Long
'
    TimerKeyboardHook.Enabled = False                                                           'Gerbing 29.03.2015
    SammelTextGetAsyncKeyState = ""                                                             'Gerbing 29.03.2015
    If gblnComefromVideo = True Then                                                            'Gerbing 29.03.2015
        If PublicPlayVideosWith = "10" Then                                                     'Gerbing 05.08.2013
            frmVideo.WMP.url = ""   ' auskommentiert Gerbing 27.11.2016                                                            'Gerbing 25.10.2013
            frmVideo.WMP.Controls.stop                                                          'Gerbing 01.09.2008
'            Form1.Show                             '29.03.2015 auskommentiert weil es sonst flackert von einen Video zum nächsten Gerbing 05.08.2013
'            frmVideo.Hide                          '29.03.2015 auskommentiert weil es sonst flackert von einen Video zum nächsten Gerbing 05.08.2013
        End If
        If Query.CheckUseAudioComments.Value = 1 Then                                           'Gerbing 26.10.2011
            frmStartSoundAutomatisch.MediaPlayer1.Controls.stop
        End If
    End If
End Sub

Private Sub TimerToPlayVideo_Timer()
    'wird benutzt um Videos mit dem internen Mediaplayer kontinuierlich abzuspielen
    TimerToPlayVideo.Enabled = False
    If F6Continous = True Then
        Unload F5MehrereZeilen                                                                  'Gerbing 15.11.2012
        'wenn blnTimer1Enabled = False dann handelt es sich um ein Video, sonst um ein Foto
        If frmVideo.WMP.playState <> 3 Then   '3=playing                                        'Gerbing 01.09.2008
            frmVideo.WMP.url = gstrFRODN    'sonst wird ein Video übersprungen wer weiß warum   'Gerbing 01.09.2008
            frmVideo.WMP.Controls.play                                                      'Gerbing 01.09.2008
        End If
        Exit Sub
    End If
End Sub

Private Function KontrolleFreischalteschlüssel(strFreischalt As String)
    'Die Gültigkeit des Freischalteschlüssels muss geprüft werden.
    'rc = 0 bei Gültigkeit
    'rc = 1 bei ungültig
    
    Dim n As Integer
    Dim i As String
    Dim lngSumme As Long
    Dim lngKontrolle As Long
    Dim lngRest As Long
    Dim lngAnzahlNumeric As Integer
    Dim Buchst As String
    Dim pos As Integer
    Dim PrüfB As String
    Dim intRest As Integer
    Dim ErsterB As Boolean
    Dim strS1 As String
    Dim strS2 As String
    Dim strS3 As String
    Dim strS4 As String
    Dim strS5 As String
    Dim strName As String
    
    Buchst = "ABCDEFGHIJKLMNOPQRSTUVWYXZ"
    strS1 = Mid(strFreischalt, 1, 5)
    strS2 = Mid(strFreischalt, 7, 5)
    strS3 = Mid(strFreischalt, 13, 5)
    strS4 = Mid(strFreischalt, 19, 5)
    strS5 = Mid(strFreischalt, 25, 5)
    If strFreischalt <> "" Then
        On Error Resume Next
        strName = Mid(strFreischalt, 31, Len(strFreischalt) - 30)
        On Error GoTo 0
    End If
    If Len(strS1) <> 5 And Len(strS2) <> 5 And Len(strS3) <> 5 And Len(strS4) <> 5 And Len(strS5) <> 5 Then
        KontrolleFreischalteschlüssel = 1             'ungültig weil keine 5 5-stelligen Kolonnen
        Exit Function
    End If
    '-------------------------------------------------------------------------------------------------
    'strS1
    lngAnzahlNumeric = 0
    For n = 1 To 5                              'enthaltenen Zahlen aufsummieren
        i = Mid(strS1, n, 1)
        If IsNumeric(i) Then
            lngAnzahlNumeric = lngAnzahlNumeric + 1
            lngSumme = lngSumme + i
        End If
    Next n
    If lngAnzahlNumeric = 5 Then                'Man darf es nicht mit lauter Zahlen probieren dürfen
        KontrolleFreischalteschlüssel = 1
        Exit Function
    End If
    '------------------------------------------------------------------------------------------------
    'strS2
    lngAnzahlNumeric = 0
    For n = 1 To 5
        i = Mid(strS2, n, 1)
        If IsNumeric(i) Then
            lngAnzahlNumeric = lngAnzahlNumeric + 1
            lngSumme = lngSumme + i
        End If
    Next n
    If lngAnzahlNumeric = 5 Then                'Man darf es nicht mit lauter Zahlen probieren dürfen
        KontrolleFreischalteschlüssel = 1
        Exit Function
    End If
    '------------------------------------------------------------------------------------------------
    'strS3
    lngAnzahlNumeric = 0
    For n = 1 To 5
        i = Mid(strS3, n, 1)
        If IsNumeric(i) Then
            lngAnzahlNumeric = lngAnzahlNumeric + 1
            lngSumme = lngSumme + i
        End If
    Next n
    If lngAnzahlNumeric = 5 Then                'Man darf es nicht mit lauter Zahlen probieren dürfen
        KontrolleFreischalteschlüssel = 1
        Exit Function
    End If
    '------------------------------------------------------------------------------------------------
    'strS4
    lngAnzahlNumeric = 0
    For n = 1 To 5
        i = Mid(strS4, n, 1)
        If IsNumeric(i) Then
            lngAnzahlNumeric = lngAnzahlNumeric + 1
            lngSumme = lngSumme + i
        End If
    Next n
    If lngAnzahlNumeric = 5 Then                'Man darf es nicht mit lauter Zahlen probieren dürfen
        KontrolleFreischalteschlüssel = 1
        Exit Function
    End If
    If lngSumme = 0 Then                        'Wenn jemand lauter Nullen probiert würde es klappen
        KontrolleFreischalteschlüssel = 1
        Exit Function
    End If
    '------------------------------------------------------------------------------------------------
    'strS5
    'die 5. Kolonne dient der Kontrolle
    For n = 1 To 5
        i = Mid(strS5, n, 1)
        If IsNumeric(i) Then
            lngKontrolle = lngKontrolle + i
        End If
    Next n
    lngRest = lngSumme Mod 7
    If lngKontrolle <> lngRest Then
        KontrolleFreischalteschlüssel = 1
        Exit Function
    End If
    '-------------------------------------------------------------------------------------------------
    'Jetzt wird geprüft,ob der Prüfbuchstabe richtig ist
    lngSumme = 0
    'strS1
    For n = 1 To 5                              'enthaltene Buchstaben aufsummieren
        i = Mid(strS1, n, 1)
        If Not IsNumeric(i) Then
            pos = InStr(1, Buchst, i)
            lngSumme = lngSumme + pos
        End If
    Next n
    '-------------------------------------------------------------------------------------------------
    'strS2
    For n = 1 To 5                              'enthaltene Buchstaben aufsummieren
        i = Mid(strS2, n, 1)
        If Not IsNumeric(i) Then
            pos = InStr(1, Buchst, i)
            lngSumme = lngSumme + pos
        End If
    Next n
    '-------------------------------------------------------------------------------------------------
    'strS3
    For n = 1 To 5                              'enthaltene Buchstaben aufsummieren
        i = Mid(strS3, n, 1)
        If Not IsNumeric(i) Then
            pos = InStr(1, Buchst, i)
            lngSumme = lngSumme + pos
        End If
    Next n
    '-------------------------------------------------------------------------------------------------
    'strS4
    For n = 1 To 5                              'enthaltene Buchstaben aufsummieren
        i = Mid(strS4, n, 1)
        If Not IsNumeric(i) Then
            pos = InStr(1, Buchst, i)
            lngSumme = lngSumme + pos
        End If
    Next n
    '-------------------------------------------------------------------------------------------------
    'die 5. Kolonne dient der Kontrolle
    ErsterB = True
    For n = 1 To 5                              'ersten Buchstaben suchen
        i = Mid(strS5, n, 1)
        If Not IsNumeric(i) Then
            If ErsterB = True Then
                PrüfB = i
            Else
                pos = InStr(1, Buchst, i)
                lngSumme = lngSumme + pos
            End If
            ErsterB = False
        End If
    Next n
    '-------------------------------------------------------------------------------------------------
    'strName                                    'Gerbing 13.10.2005
    'alle Buchstaben vom Name
    For n = 1 To Len(strName)                   'alle Buchstaben aufsummieren
        i = Mid(strName, n, 1)
        pos = InStr(1, Buchst, i)
        lngSumme = lngSumme + pos
    Next n
    '---------------------------------------------------------------------------------------------------
    'jetzt Prüfbuchstabe ausrechnen
    intRest = lngSumme Mod 26
    intRest = intRest + 1
    If PrüfB <> Mid(Buchst, intRest, 1) Then
        KontrolleFreischalteschlüssel = 1
        Exit Function
    End If
    KontrolleFreischalteschlüssel = 0           'Freischalteschlüssel ist gültig
End Function

Private Sub SpracheFestlegen()
    Dim n As Long
    Dim strTemp As String
    Dim errLoop As ADODB.Error

    On Error GoTo 0
    'Untersuche ob Access-Version oder SQL-Server-Version
    'im Fall von SQL-Server-Version wir das frmConnect Formular gezeigt
    'strTemp = Dir(gstrFotosMdbLocation & "\Fotos.mdb")
    'If strTemp <> "" Then
    If file_path_exist(gstrFotosMdbLocation & "\Fotos.mdb") = True Then
        'Access-Version wenn fotos.mdb da ist
        'strTemp = Dir(gstrFotosMdbLocation & "\$Fotos.mdb")
        'If strTemp = "" Then
        If file_path_exist(gstrFotosMdbLocation & "\$Fotos.mdb") = False Then
            'Datei nicht vorhanden
            Msg = gstrFotosMdbLocation & "\$Fotos.mdb" & vbNewLine
'            msg = msg & "Diese Datei kann nicht gefunden werden." & vbNewLine
'            msg = msg & "Sie können diese Datei selbst erzeugen durch Kopieren von fotos.mdb"
            Msg = Msg & LoadResString(2311 + Sprache) & vbNewLine
            Msg = Msg & LoadResString(2212 + Sprache)
            'MsgBox Msg
            MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
            End
        End If
    Else
        'SQL-Server-Version wenn fotos.mdbnichtda
        If gblnProversion = True Then                                                       'Gerbing 04.03.2012
            'SQL-Server-Version aber nur bei ProVersion
            gblnSQLServerConnected = False
            frmConnectSQL.Show 1
            If gblnSQLServerConnected = False Then
                'msgbox "no connection to sql server"
                MsgBox LoadResString(2460 + Sprache)
                End
            End If
        Else                                                                                'Gerbing 04.03.2012
            Msg = LoadResString(2035 + Sprache) & gstrFotosMdbLocation & "\Fotos.mdb" & LoadResString(2036 + Sprache)
            'Die Datei  ...\Fotos.mdb  ist nicht vorhanden
            'MsgBox Msg
            MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
            End
        End If
    End If
    On Error GoTo QUERYERR
    If gblnSQLServerVersion = True Then
        PublicDatagridCaption = PublicSQLServer & " " & PublicSQLDatabase
    Else
        'Set DollarDBado = New ADODB.Connection
        Set DollarDBado = CreateObject("ADODB.Connection")
        'DollarDBado.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gstrFotosMdbLocation & "\$fotos.mdb"
        'DollarDBado.ConnectionString = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & gstrFotosMdbLocation & "\$fotos.mdb"
        DollarDBado.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & gstrFotosMdbLocation & "\$fotos.mdb"
        'MessageBoxW 0, StrPtr(DollarDBado.ConnectionString), StrPtr("GERBING Fotoalbum"), vbInformation
        DollarDBado.mode = adModeReadWrite          'adModeRead=1=Read-only.    'adModeReadWrite=3=Read/write.
        'DollarDBado.Open
        DollarDBado.Open DollarDBado.ConnectionString
        
        
        'Set DBado = New ADODB.Connection
        Set DBado = CreateObject("ADODB.Connection")
        'DBado.ConnectionString = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & gstrFotosMdbLocation & "\fotos.mdb"
        'DBado.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gstrFotosMdbLocation & "\fotos.mdb"
        DBado.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & gstrFotosMdbLocation & "\fotos.mdb"
        'MessageBoxW 0, StrPtr(DBado.ConnectionString), StrPtr("GERBING Fotoalbum"), vbInformation
        DBado.mode = adModeReadWrite
        'DBado.Open
        DBado.Open DBado.ConnectionString
        PublicDatagridCaption = gstrFotosMdbLocation & "\fotos.mdb"
        Set oCat = New ADOX.Catalog                                                         'Gerbing 16.10.2014
        Set oCat.ActiveConnection = DBado                                                   'Gerbing 16.10.2014
        '------------------------------------------------------------------------------------------------------
        'Kontrolle ob die Datenbank schreibgeschützt ist                                    'Gerbing 23.11.2017
        On Error Resume Next
        SQL = "UPDATE FET SET FN = 'test'"
        Set adoRs = New ADODB.Recordset
        With adoRs
            .ActiveConnection = DBado                                             'Gerbing 23.11.2017
            .CursorType = adOpenDynamic
            '.CursorLocation = Query.enumCursorOrt
            .Source = SQL
            '     .CacheSize = 2
            .Open
        End With
        If Err.Number <> 0 Then
            gblnSchreibgeschützt = True
        Else
            gblnSchreibgeschützt = False
        End If
    End If
    If PublicLanguage = "9" Then                                                            'Gerbing 14.11.2007 'Gerbing 04.12.2011
        Call SpaltenBreiteKontrolle
        frmSprache.Show vbModal
        On Error GoTo 0
        'Wenn PublicLanguage immer noch = 9 dann konnte nicht in fotos.ini geschrieben werden
        Call GlL
        If PublicLanguage = "9" Then                                                         'Gerbing 02.09.2008 'Gerbing 04.12.2011
            Call VierUrsachenFürSchreibsperre
            End
        End If
'        SQL = "select top * from Fotos"
'        With rstsql
'            .Source = SQL
'            .ActiveConnection = DBsql
'            .CursorType = adOpenForwardOnly
'            .LockType = adLockOptimistic
'            .CursorLocation = adUseClient
'            .Open
'        End With
'
'        'untersuchen ob ein dbHyperlinkField dabei ist
'        For n = 0 To rstsql.Fields.Count - 1
'            If rstsql.Fields(n).Attributes() = 32770 Then                 '32770=dbHyperlinkField
'                'erstes Item in der Collection hat Nummer 1
'                'HyperlinkFieldColumns.Add rst2.Fields(n).Name
'                HyperlinkFieldColumns.Add n                               'beispielsweise Spalte 19
'            End If
'        Next n
'        rstsql.Close
        End
    Else
        '---------------------------------------------------------------------------------------------
        'PublicLanguage <> "9"
        SQL = "SELECT * From fotos WHERE not filename Is Null;"

        On Error Resume Next
        'On Error GoTo 0
        'On Error GoTo QUERYERR
        If rstsql Is Nothing Then
            Set rstsql = New ADODB.Recordset
        Else
            rstsql.Close
        End If
        Err.Number = 0
        With rstsql
            .Source = SQL
            .ActiveConnection = DBado                                               'Gerbing 23.11.2017
            .CursorType = adOpenForwardOnly
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
        '-------------------------------------------------------------
        If Err.Number <> 0 Then
            Call WriteGlL(0)     'Rückschreiben deutsch in fotos.ini
            Sprache = 0
            Call GlL                                                                        'Gerbing 02.09.2008
            If PublicLanguage = "1" Then                                                    'Gerbing 04.12.2011
                Call VierUrsachenFürSchreibsperre
                End
            End If
        Else
            Call WriteGlL(1)     'Rückschreiben english in fotos.ini
            Sprache = 3000
            Call GlL                                                                        'Gerbing 02.09.2008
            If PublicLanguage = "0" Then                                                    'Gerbing 04.12.2011
                Call VierUrsachenFürSchreibsperre
                End
            End If
        End If
        rstsql.Close
'        If gblnSQLServerVersion = False Then
'            'es ist keine SQL-server version - bei SQL server gibt es kein dbHyperlinkField
'            On Error GoTo 0
'            SQL = "select * from Fotos"
'            If rstsql Is Nothing Then
'                Set rstsql = New ADODB.Recordset
'            Else
'                On Error Resume Next
'                rstsql.Close
'                On Error GoTo 0
'            End If
'            With rstsql
'                .Source = SQL
'                .ActiveConnection = DBsql
'                .CursorType = adOpenForwardOnly
'                .LockType = adLockOptimistic
'                .CursorLocation = adUseClient
'                .Open
'            End With
'            'untersuchen ob ein dbHyperlinkField dabei ist
'            For n = 0 To rstsql.Fields.Count - 1
'                If rstsql.Fields(n).Attributes() = 32770 Then                 '32770=dbHyperlinkField
'                    'erstes Item in der Collection hat Nummer 1
'                    'HyperlinkFieldColumns.Add rst2.Fields(n).Name
'                    HyperlinkFieldColumns.Add n                               'beispielsweise Spalte 19
'                End If
'            Next n
'            rstsql.Close
'        End If
    End If
    Exit Sub
QUERYERR:
    If DBado.Errors.Count > 0 Then
        For Each errLoop In DBado.Errors
            Msg = "Fehler Nr.: " & errLoop.Number & vbCr & errLoop.Description & " " & DBado.ConnectionString   'Gerbing 22.11.2017
            MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
        Next errLoop
    End If
    Resume Next
End Sub

Private Sub EinBRückADO()
    On Error Resume Next
    frmGridAndThumb.Adodc1.Recordset.MovePrevious
    'If ERR.Number <> 0 Then
        If frmGridAndThumb.Adodc1.Recordset.BOF Then         'ich bin schon am Anfang
            On Error GoTo 0
            frmGridAndThumb.Adodc1.Recordset.MoveLast
        Else
            If Err.Number = 0 Then
                On Error GoTo 0                                                     'Gerbing 29.12.2005
                Exit Sub
            End If
            If Err = 91 Then    'Objektvariable oder With-Blockvariable nicht festgelegt
                Msg = "Es wurde kein einziger Datensatz gefunden." & NL
                Msg = Msg & "Mit der F8-Taste können Sie die Suche wiederholen"
                'MsgBox msg                 'Gerbing 08.11.2005
                MsgBox LoadResString(2007 + Sprache) & NL & LoadResString(2008 + Sprache)
                On Error GoTo 0                                                     'Gerbing 29.12.2005
                Exit Sub
            Else
                Msg = "Errortext=" & Err.Description & NL
                Msg = Msg & "Errorcode=" & Err.Number & NL & NL
                MsgBox Msg
            End If
        End If
    'End If
    On Error GoTo 0
End Sub


Private Sub EinBVorADO()
    Dim DateinamenErweiterung As String
    Dim sngWert As Single                                                               'Gerbing 22.11.2010
    Dim lngWert As Long
    Dim Obergrenze As Long
    Dim Untergrenze As Long

    
Anfang:
    On Error Resume Next
    If F6Continous = True Then                                                          'Gerbing 22.11.2010
        If gblnWertxFormOptZufall = True Then
            Obergrenze = frmGridAndThumb.Adodc1.Recordset.RecordCount
            Untergrenze = 0
            Randomize
            sngWert = (Obergrenze - Untergrenze + 1) * Rnd + Untergrenze
            lngWert = Round(sngWert)
            If lngWert > Obergrenze Then
                lngWert = Obergrenze - 1
            End If
            frmGridAndThumb.Adodc1.Recordset.Move lngWert, adBookmarkFirst
        Else
            frmGridAndThumb.Adodc1.Recordset.MoveNext
        End If
    Else
        frmGridAndThumb.Adodc1.Recordset.MoveNext
    End If
    'If ERR.Number <> 0 Then
        If frmGridAndThumb.Adodc1.Recordset.EOF Then         'ich bin schon am Ende
            On Error GoTo 0
            frmGridAndThumb.Adodc1.Recordset.MoveFirst
        Else
            If Err.Number <> 0 Then
                If Err = 91 Then    'Objektvariable oder With-Blockvariable nicht festgelegt
                    Msg = "Es wurde kein einziger Datensatz gefunden." & NL
                    Msg = Msg & "Mit der F8-Taste können Sie die Suche wiederholen"
                    'MsgBox msg                 'Gerbing 08.11.2005
                    MsgBox LoadResString(2007 + Sprache) & NL & LoadResString(2008 + Sprache)
                    On Error GoTo 0                                                     'Gerbing 29.12.2005
                    Exit Sub
                Else
                    Msg = "Errortext=" & Err.Description & NL
                    Msg = Msg & "Errorcode=" & Err.Number & NL & NL
                    If Err.Number <> -2147217842 And Err.Number <> 481 Then             'Gerbing 29.03.2012
                        MsgBox Msg
                    End If
                End If
            End If
        End If
    'End If
    On Error GoTo 0
    'Ich will für alle Link-Filetypes die Automatic ausser Kraft setzen     'Gerbing 29.12.2005
    'd.h. wenn ich mit MoveNext auf eine Datei mit Link-Filetyp stosse,
    'dann ignoriere ich diese Datei und gehe zur nächsten
    If F6Continous = True Then
        Call FRODateiname
        DateinamenErweiterung = Right(gstrFRODN, 3)
        DateinamenErweiterung = UCase(DateinamenErweiterung)
        Select Case DateinamenErweiterung
        Case "BMP", "DIB", "EMF", "GIF", "ICO", "JPG", "PNG", "TIF", "TIFF", "WMF"
            'das sind die native mode formats für fotos
        Case "AVI", "MPG", "PEG", "MOV", "MPE", "ASF", "ASX", "WMV", "MP4", "MKV", "FLV"      'Gerbing 10.12.2017
            'das sind die native mode formats für videos
            If StrComp(Right(gstrFRODN, 4), "JPEG", vbTextCompare) = 0 Then     'Gerbing 21.08.2006
                DoEvents
                GoTo Anfang             'gehe zur nächsten Datei
            End If
        Case Else
            'das sind die link mode formats
            DoEvents
            GoTo Anfang             'gehe zur nächsten Datei
        End Select
    End If
End Sub
Public Sub AbrufenBildPosList()                          'Gerbing 29.07.2006
    Dim i As Long
    Dim pos As Long
    
    'Wenn gblnCheckSpeichernBildPosition eingeschaltet ist, wird der zum aktuellen Datensatz gehörende
    'Dateiname im Array gesucht und wenn vorhanden
    'zur Bildpositionierung benutzt
    'aber nicht, wenn alle 4 Einträge Null sind
    
    If gblnCheckSpeichernBildPosition = True Then
'        If Query.RecordCount > 99 Then                                         'Gerbing 04.10.2007 26.09.2014
'            Exit Sub
'        End If
        For i = 0 To Query.RecordCount - 1
            pos = InStr(1, BildPosList(i).Dateiname, gstrFRODN, vbTextCompare)
            If pos <> 0 Then
            glngDiffX = BildPosList(i).Left
            glngDiffY = BildPosList(i).Top
            glngZoomProzent = BildPosList(i).ZoomPercent
                Exit For
            End If
        Next i
    End If
End Sub

Public Sub SpeichernInBildPosList()
    Dim i As Long
    Dim pos As Long

    'Wenn gblnCheckSpeichernBildPosition eingeschaltet ist, wird der zum aktuellen Datensatz gehörende
    'Dateiname im Array gesucht und wenn vorhanden die 4 Bildpositionen überschrieben
    'wenn nicht vorhanden hinten angefügt
    
    If gblnCheckSpeichernBildPosition = True Then
'        If Query.RecordCount > 99 Then                                          'Gerbing 04.10.2007 26.09.2014
'            Exit Sub
'        End If
        For i = 0 To Query.RecordCount - 1
            If Mid(BildPosList(i).Dateiname, 1, 1) = "?" Then
                Exit For
            End If
            pos = InStr(1, BildPosList(i).Dateiname, gstrFRODN, vbTextCompare)
            If pos <> 0 Then
                BildPosList(i).Left = glngDiffX
                BildPosList(i).Top = glngDiffY
                BildPosList(i).ZoomPercent = glngZoomProzent
                Exit For
            End If
        Next i
        BildPosList(i).Dateiname = gstrFRODN
        BildPosList(i).Left = glngDiffX
        BildPosList(i).Top = glngDiffY
        BildPosList(i).ZoomPercent = glngZoomProzent
    End If
End Sub

Private Sub TimerVideoDuration_Timer()
    'zur Steuerung des externen mediaplayers wenn dieser kontinuierlich abspielen soll
    
    'Wenn glngVideoDuration ungleich Null ist und If F6Continous = True,
    'muss das nächste Video gestartet werden
    'TimerVideoDuration.Interval = 1000 darf nicht verändert werden
    'In glngVideoDuration steht beispielsweise 55 bei 55 Sekunden

    If glngVideoDuration <> 0 Then
        glngVideoDuration = glngVideoDuration - 1
        If glngVideoDuration <= 0 Then
            TimerVideoDuration.Enabled = False                                                  'Gerbing 19.10.2007
            If F6Continous = True Then
                Unload F5MehrereZeilen                                                          'Gerbing 15.11.2012
                Call GeheEinBildVorwärts
                Call BildAnzeigen
            End If
        End If
    End If
End Sub

Private Sub FremdeFotosMdb()                                                                    'Gerbing 07.11.2011
    Dim Msg As String
    
    Me.Caption = "GERBING Fotoalbum"                                                            'Gerbing 08.05.2020
    Me.Show

Begin:
    '(ByVal Filter$, ByVal InitialDir$, ByVal Title$) as String
    gstrNetzwerkDir = ShowOpenUnicodeFotosMdb(Me)    '2458=Standort der fotos.mdb
    AppActivate Me.Caption                                                                      'Gerbing 08.05.2020
    
    'Convert the file name to be used
    gstrNetzwerkDir = ConvertFileName(gstrNetzwerkDir)
    If gstrNetzwerkDir = "" Then
        gstrFotosMdbLocation = AppPath
        Exit Sub
    End If
    If Mid(gstrNetzwerkDir, Len(gstrNetzwerkDir) - 9, 1) <> "\" Then
        Msg = LoadResString(2459 + Sprache)                  '2459=Sie müssen die Datei fotos.mdb auswählen
        MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
        GoTo Begin
    End If
    If StrComp(Right(gstrNetzwerkDir, 9), "fotos.mdb", vbTextCompare) = 0 Then
        gstrFotosMdbLocation = Mid(gstrNetzwerkDir, 1, Len(gstrNetzwerkDir) - 10)
    Else
        Msg = LoadResString(2459 + Sprache)                 '2459=Sie müssen die Datei fotos.mdb auswählen
        MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
        GoTo Begin
    End If
End Sub

Public Sub KontrolleManagement()
    'Wenn in Tabelle LoggedInUsers Spalte Management nicht 'IN ...' entdeckt wird, dann hat der SQL Server Administrator Reset für den
    'betreffenden user gemacht, oder 'Reset all'
    'Wenn die Uhrzeit verpfuscht ist, dann hat jemand an den Einträgen der Tabelle LoggedInUsers Spalte Management herumgepfuscht
    'um meine Lizenzenkontrolle zu überlisten
    
    Dim strManagement As String
    Dim strUsername As String
    Dim ManagementUnkodiert As String
    Dim KontrolleDate As Date
    
    If gstrAllowedlicenses = 99 Then Exit Sub
    Set rstsql = New ADODB.Recordset
    With rstsql
        .Source = "select * from loggedinusers where (username = N'" & gstrLoggedInName & "')"  'Gerbing 04.03.2013 N voranstellen wegen unicode
        .ActiveConnection = DBado
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    strManagement = rstsql.Fields("Management")
    strUsername = rstsql.Fields("username")
    rstsql.Close
    ManagementUnkodiert = Crypt(strManagement, strUsername, False)
    If Mid(ManagementUnkodiert, 1, 3) <> "IN " Then
        'MsgBox "Der sql server administrator hat Sie ausgeloggt. Das Programm wird beendet."
        MsgBox LoadResString(1818 + Sprache)
        Call Query.Beenden
        End
    End If
    On Error Resume Next
    KontrolleDate = Mid(ManagementUnkodiert, 4, Len(ManagementUnkodiert) - 3)
    If Err.Number <> 0 Then
        'MsgBox "Unerlaubter Eingriff in die Tabelle LoggedInUsers. Das Programm wird beendet."
        MsgBox LoadResString(1819 + Sprache)
        Call Query.Beenden
        End
    End If
    
End Sub

Public Sub FelderAusfüllenF5MehrereZeilen()
    F5MehrereZeilen.chkIptc_Click
End Sub

Private Sub ErsatzFürAppActivate(Klasse, AppName)
    Dim WndHnd As Long
    Dim rc As Long

    WndHnd = FindWindow(Klasse, AppName)
    If WndHnd = 0 Then
        Msg = "FindWindow konnte " & AppName & " nicht finden"
        MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
        Exit Sub
    End If
    rc = SetForegroundWindow(WndHnd)
    rc = SetActiveWindow(WndHnd)
End Sub

Public Function EXIFLesenOrientation(Filename As String)                                                              'Gerbing 10.05.2019
    Dim EXIFInfo As String
    
    If StrComp(Right(Filename, 3), "JPG", vbTextCompare) <> 0 Then
        Exit Function
    End If
    
    Form1.EXF.ImageFile = Filename 'set the image file property, read metainfo, parse metainfo
    '
    'EXF.ListInfo ist ein String mit vbCrLf
    '
    EXIFInfo = Form1.EXF.ListInfo 'list all tags into the text box
End Function

Public Sub MyDrawImage(Filename As String, ZoomPercent As Long)
    Dim SaveMyZoomPercent As Long                                                                   'Gerbing 10.05.2019
    
    gstrEXIFOrientation = ""                                                                        'Gerbing 10.05.2019
    Call EXIFLesenOrientation(Filename)                                                             'Gerbing 10.05.2019
    
    'Weil bei Verzweigung zu RenamMdb das aktuell gezeigte Bild nicht gelöscht oder umgenannt werden kann, 'Gerbing 25.03.2018
    'ist es möglich, dass m_lngInstance mit gdiPlusShutdown entladen war
    'und ich muss zuerst GdiplusStartup ausführen
    udtData.GdiplusVersion = 1                                          'Gerbing 25.03.2018
    If GdiplusStartup(m_lngInstance, udtData, 0) Then
        MsgBox "GDI+ could not be initialized", vbCritical
        Exit Sub
    End If                                                              'Gerbing 25.03.2018
    Picture1.Picture = LoadPicture("")
    
    If GdipCreateFromHDC(Form1.Picture1.hDC, m_lngGraphics) Then        'Gerbing 31.12.2012
        MsgBox "Graphics object could not be created", vbCritical
        Exit Sub
    End If
    If gblnBildBeschreibung = False Then
        txtBildbeschreibung.Visible = False
        Picture1.Top = 0                                                'Gerbing 11.08.2012
    Else
        Picture1.Top = txtBildbeschreibung.height                       'Gerbing 11.08.2012
        txtBildbeschreibung.Text = Filename
        txtBildbeschreibung.width = Me.width
        txtBildbeschreibung.Visible = True
    End If
    
    MyZoomPercent = ZoomPercent
    retcode = GdipLoadImageFromFile(StrPtr(Filename), lngPointer)
    If retcode <> 0 Then
        MsgBox "Fehler bei GdipLoadImageFromFile"                       'Gerbing 29.04.2018
    End If
    
    If gstrEXIFOrientation = "8" Then GdipImageRotateFlip lngPointer, Rotate270FlipNone             'Rotate270FlipNone = 3  Gerbing 30.09.2019
    If gstrEXIFOrientation = "7" Then GdipImageRotateFlip lngPointer, Rotate270FlipX                'Rotate270FlipX = 7     Gerbing 30.09.2019
    If gstrEXIFOrientation = "6" Then GdipImageRotateFlip lngPointer, Rotate90FlipNone              'Rotate90FlipNone = 1   Gerbing 30.09.2019
    If gstrEXIFOrientation = "5" Then GdipImageRotateFlip lngPointer, Rotate90FlipX                 'Rotate90FlipX = 5      Gerbing 30.09.2019
    If gstrEXIFOrientation = "4" Then GdipImageRotateFlip lngPointer, Rotate180FlipX                'Rotate180FlipX = 6     Gerbing 30.09.2019
    If gstrEXIFOrientation = "3" Then GdipImageRotateFlip lngPointer, Rotate180FlipNone             'Rotate180FlipNone = 2  Gerbing 30.09.2019
    If gstrEXIFOrientation = "2" Then GdipImageRotateFlip lngPointer, RotateNoneFlipX               'RotateNoneFlipX = 4    Gerbing 30.09.2019
    
    retcode = GdipGetImageDimension(lngPointer, sngWidth, sngHeight)
    If retcode <> 0 Then
        MsgBox "Fehler bei GdipGetImageDimension"                       'Gerbing 29.04.2018
    End If
    'X und Y ausrechnen
    On Error Resume Next                                                'Gerbing 05.09.2012
    If Query.chkFensterGrößeÄnderbar.Value = 1 Then
        X = 0
        Y = 0
    Else
        X = (screenWidth \ 2) - ((ZoomPercent / 100) * (sngWidth \ 2))
        Y = (screenHeight \ 2) - ((ZoomPercent / 100) * (sngHeight \ 2))
    End If
    On Error GoTo 0                                                     'Gerbing 05.09.2012
    X = X + glngDiffX
    Y = Y + glngDiffY
    If gblnRechteckLupeScharf = True Then
        'Bei Rechteck-Zoom ist uninteressant was als X oder Y bisher errechnet wurde
        X = DifferenzX * gdblZoomFaktor
        Y = DifferenzY * gdblZoomFaktor
        MyZoomPercent = MyZoomPercent * gdblZoomFaktor
        SaveMyZoomPercent = MyZoomPercent                                                       'Gerbing 10.05.2019
        Shape1.Visible = False
        glngDiffX = 0
        glngDiffY = 0
        glngZoomProzent = 100
        '-----------------------------------------------------------------------------
        'anschließend muß der Mauszeiger wieder zurückgesetzt werden    'Gerbing 27.11.2012
        If gblnMouseSichtbar = True Then
            If gblnRechteckLupeScharf = True Then                                               'Gerbing 29.07.2007
                Me.MousePointer = vbCustom                                                      'Gerbing 29.07.2007
                'Me.MouseIcon = LoadPicture(AppPath & "\SquareZoom.ico")                        'Gerbing 29.07.2007
                Me.MouseIcon = LoadResPicture(105, 1)                                           'Gerbing 04.03.2013
            Else
                Me.MousePointer = vbDefault                                                     'Gerbing 29.07.2007
            End If
        Else
            Me.MousePointer = vbCustom      '99                                                 'Gerbing 29.07.2007
            'Me.MouseIcon = LoadPicture(AppPath & "\MOUSE01.ICO")                               'Gerbing 29.07.2007
            Me.MouseIcon = LoadResPicture(104, 1)                                               'Gerbing 04.03.2013
        End If
        gblnComeFromRechtecklupe = True                                                         'Gerbing 14.10.2014
    End If
    glngSaveX = X
    glngSaveY = Y
    ' Setzen der Optimierungsmodis
    retcode = GdipSetSmoothingMode(m_lngGraphics, SmoothingModeNone)
    If retcode <> 0 Then
        MsgBox "Fehler bei GdipSetSmoothingMode"                                                'Gerbing 29.04.2018
    End If
    retcode = GdipSetInterpolationMode(m_lngGraphics, InterpolationModeHighQualityBicubic)
    If retcode <> 0 Then
        MsgBox "Fehler bei GdipSetInterpolationMode"                                            'Gerbing 29.04.2018
    End If
    retcode = GdipSetPixelOffsetMode(m_lngGraphics, PixelOffsetModeNone)
    If retcode <> 0 Then
        MsgBox "Fehler bei GdipSetPixelOffsetMode"                                              'Gerbing 29.04.2018
    End If
    retcode = GdipSetCompositingQuality(m_lngGraphics, CompositingQualityDefault)
    If retcode <> 0 Then
        MsgBox "Fehler bei GdipSetCompositingQuality"                                           'Gerbing 29.04.2018
    End If
    retcode = GdipSetCompositingMode(m_lngGraphics, CompositingModeSourceOver)
    If retcode <> 0 Then
        MsgBox "Fehler bei GdipSetCompositingMode"                                              'Gerbing 29.04.2018
    End If
    ' zoomen
    retcode = GdipDrawImageRect(m_lngGraphics, lngPointer, X, Y, sngWidth * MyZoomPercent \ 100, sngHeight * MyZoomPercent \ 100)
    If retcode <> 0 Then
        MsgBox "Fehler bei GdipDrawImageRect"                                                   'Gerbing 29.04.2018
    End If
    gblnRechteckLupeScharf = False      'Der Rechteck-Zoom funktioniert nur einmal                  'Gerbing 14.10.2014 10.05.2019
    '----------------------------------------------------------------------------------------------------------------------------
    'oh Wunder das Bild bleibt sichtbar                                                         '25.03.2018
    If lngPointer Then
        retcode = GdipDisposeImage(lngPointer)
        lngPointer = 0
    End If
    If m_lngGraphics Then
        If GdipDeleteGraphics(m_lngGraphics) Then
            MsgBox "Graphics object could not be deleted", vbCritical                           'Gerbing 29.04.2018 auskommentiert
        End If
    End If
    GdiplusShutdown m_lngInstance                                                               '25.03.2018
    '---------------------------------------------------------------------------------------------------------------------------
    Me.Refresh
    Form1.Picture1.Refresh                                                                      'Gerbing 10.05.2019
    'Me.Show führt zum Flackern wenn bei sichtbarer Form frmGridAndThumb Bildwechsel gemacht wird, egal ob durch Klick
    'auf eine Zeile im Grid oder durch Klick auf ein Thumbnail
    If gblnComeFromThumbs = False Then                                                          'Gerbing 29.03.2015
        Me.Show
    End If
    On Error GoTo 0
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = vbRightButton And gblnMouseIconSquare = True Then                    'Gerbing 23.10.2014 25.04.2018
'        Exit Sub
'    End If
    If gblnRechteckLupeScharf = False And gblnfrmGridAndThumbDblClick = False Then       'Gerbing 05.05.2005
        If gblnRechteckLupeScharf = False Then
            'Wenn die Rechteck-Lupe nicht scharf ist und auch kein Doppelklick vorliegt, wird mit der linken Maustaste das Bild verschoben
            'und zwar exakt an die mit der Maus angesteuerte Stelle
            If blnComeFromError = True Then Exit Sub                                'Gerbing 15.02.2014
            If Button = vbLeftButton Then
                StartX = X
                StartY = Y
                Me.MousePointer = vbCustom
                If gblnCheckSpeichernBildPosition = True Then
                    'Me.MouseIcon = LoadPicture(AppPath & "\FourArrowsSave.ico")    'Gerbing 29.07.2007
                    Me.MouseIcon = LoadResPicture(102, 1)                           'Gerbing 04.03.2013
                Else
                    'Me.MouseIcon = LoadPicture(AppPath & "\FourArrows.ico")        'Gerbing 29.07.2007
                    Me.MouseIcon = LoadResPicture(101, 1)                           'Gerbing 04.03.2013
                End If
                Exit Sub
            End If
        End If
    End If
    If Button = vbLeftButton Then
        'Auf dem Klickpunkt beginnt das Rechteck, anfangs mit Width und Height = 0
        'Me.Cls
        Shape1.width = 0
        Shape1.height = 0
        StartX = X
        StartY = Y
        Shape1.Left = X
        Shape1.Top = Y
        EndX = X
        EndY = Y
        Shape1.Visible = True
    End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, MyX As Single, MyY As Single)
    Dim brush As Long

'    If Button = vbRightButton And gblnMouseIconSquare = True Then        'Gerbing 23.10.2014 25.04.2018
'        Exit Sub
'    End If
    If Button = vbLeftButton And gblnRechteckLupeScharf = True Then
        'es wird ein Rechteck gezeichnet, das stets das Breiten/Höhen-Verhältnis des Bildschirms SGVH einhält 'Gerbing 03.09.2010
        'es wird nur gezeichnet, wenn der Nutzer von links nach rechts und von oben nach unten zieht
        RX = MyX
        If MyX > EndX And MyY > EndY Then
            If ((MyX - StartX) * 4) <> ((MyY - StartY) * 3) Then
                RX = MyX + ((MyY - StartY) * Form1.SGVH)
            End If
            If RX - StartX < 0 Then Exit Sub
            Shape1.width = RX - StartX
            Shape1.height = Shape1.width / Form1.SGVH
        End If
        EndX = MyX
        EndY = MyY
        'Debug.Print "startx= " & StartX & "/StartY= " & StartY & "/endx= " & EndX & "/endy= " & EndY
    End If
    If MyZoomPercent <> 0 Then
        XYPos.lblXPos = (MyX / Screen.TwipsPerPixelX - X) * 100 \ MyZoomPercent 'Gerbing 11.03.2017 bei / entsteht eine Zahl mit Komma
        If XYPos.lblXPos < 0 Or XYPos.lblXPos > gsngPicWidth Then
            XYPos.lblXPos = ""
        End If
        XYPos.lblYPos = (MyY / Screen.TwipsPerPixelY - Y) * 100 \ MyZoomPercent 'Gerbing 11.03.2017 bei / entsteht eine Zahl mit Komma
        If XYPos.lblYPos < 0 Or XYPos.lblYPos > gsngPicHeight Then
            XYPos.lblYPos = ""
        End If
        XYPos.lblBildgröße = gsngPicWidth & " x " & gsngPicHeight
    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim KeyCode As Integer
    Dim Msg As String
    
'    If Button = vbRightButton And gblnMouseIconSquare = True Then                   'Gerbing 23.10.2014 25.04.2018
'        Exit Sub
'    End If
    If gblnRechteckLupeScharf = False And gblnfrmGridAndThumbDblClick = False Then       'Gerbing 05.05.2005
        'Wenn die Rechteck-Lupe nicht scharf ist und auch kein Doppelklick vorliegt, wird mit der linken Maustaste das Bild verschoben
        'und zwar exakt an die mit der Maus angesteuerte Stelle
        If Button = vbLeftButton And gblnRechteckLupeScharf = False Then
            'aber nur wenn überhaupt eine Verschiebung stattgefunden hat
            If StartX <> X And StartY <> Y Then
                glngDiffX = glngDiffX + ((X - StartX) / Screen.TwipsPerPixelX)      'Gerbing 10.06.2012
                glngDiffY = glngDiffY + ((Y - StartY) / Screen.TwipsPerPixelY)      'Gerbing 10.06.2012
                Call Form1.SpeichernInBildPosList
                Me.MousePointer = vbDefault
                KeyCode = vbKeySeparator                                    'vbKeySeparator als Kennzeichnung für Verschiebung
                Shift = 0
                Call Form1.Form_KeyDown(KeyCode, Shift)
                Exit Sub                                                            'Gerbing 23.10.2014
            End If
        End If
    End If
    gblnfrmGridAndThumbDblClick = False
    
    'If Button = vbRightButton And Shape1.Visible = False Then
    If Button = vbRightButton Then
        'Me.WindowState = 2  '2=maximized
        blnHilfeboxStehenLassen = True                                              'Gerbing 08.10.2014
        Call Form1.Hilfebox
        Exit Sub
    End If

    If Shape1.Visible = False Then Exit Sub
    
    If gblnRechteckLupeScharf = True Then
        If EndX = StartX Or EndY = StartY Then
            Shape1.Visible = False
            'Wenn der Nutzer nur ins Bild klickt, bekommt er eine MsgBox
            'msg = "Wenn Sie die Rechteck-Lupe benutzen wollen, müssen Sie mit der Maus ein Rechteck zeichnen." & vbNewLine
            Msg = LoadResString(2049 + Sprache) & vbNewLine
            'msg = msg & "Sie haben nur in das Bild geklickt." & vbNewLine & vbNewLine
            Msg = Msg & LoadResString(2050 + Sprache) & vbNewLine & vbNewLine
            MsgBox Msg
            'ImageForm.Image1.Refresh                          'Gerbing 06.06.2003
            Exit Sub
        End If
        'wenn nicht ein Rechteck mit einer Mindestgröße vorliegt, wird Zoom nicht gemacht
        If (EndX - StartX) < 20 Or (EndY - StartY) < 15 Then
            'msg = "Die Rechteck-Lupe ist scharf." & vbNewLine
            Msg = LoadResString(2052 + Sprache) & vbNewLine
            'msg = msg & "Zu kleine Bildausschnitte werden nicht gezoomt." & vbNewLine
            Msg = Msg & LoadResString(2053 + Sprache) & vbNewLine
            'msg = msg & "Möglicherweise haben sie nur ins Bild geklickt" & vbNewLine
            Msg = Msg & LoadResString(2054 + Sprache) & vbNewLine
            'msg = msg & "oder nicht von links nach rechts und von oben nach unten gezogen."
            Msg = Msg & LoadResString(2055 + Sprache)
            MsgBox Msg
            Exit Sub
        End If
        DifferenzX = glngSaveX - (StartX / Screen.TwipsPerPixelX)                       'Gerbing 10.06.2012
        DifferenzY = glngSaveY - (StartY / Screen.TwipsPerPixelY)                       'Gerbing 10.06.2012
        gdblZoomFaktor = screenWidth / (Shape1.width / Screen.TwipsPerPixelX)           'Gerbing 10.06.2012
        KeyCode = vbKeySeparator                                    'vbKeySeparator als Kennzeichnung für Rechtecklupe
        Shift = 0
        Call Form1.Form_KeyDown(KeyCode, Shift)
    End If
End Sub

Private Function CompKey(KCode As Long, KText As String) As Boolean
    'Damit kann ich reagieren wenn bei einem Video F8 oder Umsch + F5 gedrückt wird
    Dim result As Integer
    Dim KeyCode As Integer
    Dim Shift As Integer
    Dim pos As Long
    
    result = GetAsyncKeyState(KCode)
    If result = -32767 Then
        SammelTextGetAsyncKeyState = SammelTextGetAsyncKeyState & KText
        CompKey = True
        If KText = "NUM G" Then
            TimerKeyboardHook.Enabled = False
            KeyCode = vbKeyF8                                               'F8
            Shift = 0
            Call Form1.Form_KeyDown(KeyCode, Shift)
        End If
        pos = InStr(1, SammelTextGetAsyncKeyState, "SHIFTNUM D", vbTextCompare)
        If pos <> 0 Then
            SammelTextGetAsyncKeyState = ""
            KeyCode = vbKeyF5
            Shift = vbShiftMask                                             'Umsch + F5 gleichzeitig
            Call Form1.Form_KeyDown(KeyCode, Shift)
        End If
    Else
        CompKey = False
    End If
End Function

Private Sub SpaltenBreiteKontrolle()                                                    'Gerbing 19.04.2015
    'Wenn die SummeSpaltenbreite höher als 15000 ist vermute ich, daß mit Twips gerechnet wurde
    'und ändere die Standardspaltenbreite auf 100 Pixel
    Dim SQL As String
    Dim SummeSpaltenbreite As Long
    Dim n As Long
    
    SQL = "SELECT * From Spaltenbreite;"

    'On Error Resume Next
    On Error GoTo 0
    'On Error GoTo QUERYERR
    If rstsql Is Nothing Then
        Set rstsql = New ADODB.Recordset
    Else
        rstsql.Close
    End If
    Err.Number = 0
    With rstsql
        .Source = SQL
        .ActiveConnection = DBado                                               'Gerbing 23.11.2017
        .CursorType = adOpenForwardOnly
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    '-------------------------------------------------------------
    If Err.Number = 0 Then
        Do Until rstsql.EOF
            SummeSpaltenbreite = SummeSpaltenbreite + rstsql.Fields("Spaltenbreite")
            rstsql.MoveNext
        Loop
        If SummeSpaltenbreite > 15000 Then
            SQL = "UPDATE SpaltenBreite SET SpaltenBreite.Spaltenbreite = 100;"
            If rstsql Is Nothing Then
                Set rstsql = New ADODB.Recordset
            Else
                rstsql.Close
            End If
            Err.Number = 0
            With rstsql
                .Source = SQL
                .ActiveConnection = DBado                                               'Gerbing 23.11.2017
                .CursorType = adOpenForwardOnly
                .LockType = adLockOptimistic
                .CursorLocation = adUseClient
                .Open
            End With
        End If
    End If
    rstsql.Close
End Sub

Private Sub KommentarNachBildanzeigen()
    Dim strTemp As String
    Dim strKommentar As String
    Dim SQL As String
    Dim MyRecordset As ADODB.Recordset                                                          'Gerbing 29.12.2011
    
    If KommentarFensterEinblenden = True Then                               'Gerbing 04.03.2013
        If gblnWasOptThumbClick = False Then                                                'Gerbing 04.05.2015
            If Not IsNull(frmGridAndThumb.Adodc1.Recordset(LoadResString(1030 + Sprache))) And frmGridAndThumb.Adodc1.Recordset(LoadResString(1030 + Sprache)) <> "" Then
                KommentarForm.Show
            End If
        Else
            gblnWasOptThumbClick = False
            'in gstrFRODN steht der Dateiname des zu untersuchenden Records, dort will ich den Kommentar entnehmen      'Gerbing 04.05.2015
            ' Recordset erstellen und öffnen adOpenStatic
            strTemp = Replace(gstrFRODN, gstrFotosMdbLocation & "\", "+:\")
            strTemp = Replace(strTemp, "'", "''")                                           'Gerbing 23.01.2018
            'Bei Dateinamen mit Hochkomma bringt Open Recordset Laufzeitfehler -> ersetzen durch 2 Hochkommas  'Gerbing 23.01.2018
            Set MyRecordset = New ADODB.Recordset
            'SQL = "SELECT Dateiname, Kommentar From Fotos Where Dateiname = " & strTemp
            SQL = "SELECT " & LoadResString(1028 + Sprache) & "," & LoadResString(1030 + Sprache) & " From Fotos Where " & LoadResString(1028 + Sprache) & "='" & strTemp & "'"
            With MyRecordset
                .Source = SQL
                .ActiveConnection = DBado                                               'Gerbing 23.11.2017
                .CursorType = adOpenStatic
                .LockType = adLockOptimistic
                .CursorLocation = adUseClient
                .Open
            End With
            If Not IsNull(MyRecordset(LoadResString(1030 + Sprache))) Then
                strKommentar = MyRecordset(LoadResString(1030 + Sprache))
                If strKommentar <> "" Then
                    KommentarForm.txtKommentar.Text = strKommentar
                    KommentarForm.Show
                End If
            End If
            MyRecordset.Close
        End If
    End If
End Sub

Public Sub EmailMitAnhangSenden()                                          'Gerbing 13.08.2017
    ' Outlook Applikation
    'erfordert Projekt -> Verweise -> Microsoft Outlook 14.0 Object Library(msoutl.olb)
    Dim ool As Outlook.Application
    Dim oInspector As Outlook.Inspector
    Dim oMail As Outlook.MailItem
    Dim myattachments As Variant
    
    ' Für Inputbox "EMailadresse-Änderung"
    Dim Mldg, Titel, Voreinstellung, MailAdress
    
    
    '     ' Adresse anzeigen und Änderung ermöglichen
    '     Mldg = "Ist die angegebene Emailadresse richtig?"
    '     Titel = "Mailadresse"
    '     Voreinstellung = "trallala@gmx.de"
    '     MailAdress = InputBox(Mldg, Titel, Voreinstellung)
    MailAdress = LoadResString(1239 + Sprache)                              '"Bitte eingeben"
    
    ' Wurde Abbrechen gedrückt, dann alles beenden
    If MailAdress = "" Then Exit Sub
    
    ' Verweis zu Outlook + neue Nachricht
    On Error GoTo Outlookfehlt
    Set ool = CreateObject("Outlook.Application")
    Set oMail = ool.CreateItem(olMailItem)
    Set myattachments = oMail.Attachments
    
    ' Befreff-Zeile
    oMail.Subject = LoadResString(1239 + Sprache)                           '"Bitte eingeben"
    
    ' An-Zeile (Empfänger)
    oMail.To = MailAdress
    'oMail.Recipients.ResolveAll                      'hier kommt error 427
    oMail.display
    
    ' Texteingabe (Nachricht selbst)
    oMail.Body = LoadResString(1239 + Sprache)                              '"Bitte eingeben"
    
    
    ' Anhang
    ' Nachfolgend ein Beispiel. Suchen Sie sich eine Datei auf
    ' Ihrem Rechner aus - vollständiger Pfad muß mitangegeben
    ' sein.
    ' Es können auch weitere Dateien angegeben werden.
    ' Hierzu einfach mit myattachments.Add "???" fortsetzen.
    myattachments.Add "" & gstrFRODN & ""
    
    ' Speicher freigeben
    Set ool = Nothing
    Set oInspector = Nothing
    Set oMail = Nothing
    Exit Sub
Outlookfehlt:
    MsgBox LoadResString(1240 + Sprache)                                    '"Outlook ist nicht installiert"
End Sub

Public Sub LöschenInDatenbankUndStandort(Dateiname As String, rst1 As ADODB.Recordset)
    Dim antwort As Long
    Dim strTemp As String
    Dim DateinameFoto As String
    Dim temp As String
    Dim temp1 As String
    Dim rc As Boolean
    
    'On Error Resume Next
    strTemp = Replace(Dateiname, "+:\", gstrFotosMdbLocation & "\")
    'Falls es einen zugehörigen Audio-Kommentar gibt, wird dieser zuerst gelöscht    'Gerbing 12.04.2006
    DateinameFoto = ErmittleDateiname(strTemp)
    If file_path_exist(DateinameFoto & ".mp3") = True Then
        temp = DateinameFoto & ".mp3"
    End If
    If file_path_exist(DateinameFoto & ".wav") = True Then
        temp = DateinameFoto & ".wav"
    End If
    If temp <> "" Then                                                              'Gerbing 04.09.2013
        rc = file_delete(temp, , True)                                              'Gerbing 04.09.2013
    End If
    rc = file_delete(strTemp, , True)                                       'Gerbing 04.09.2013
    If rc = False Then Exit Sub
    '----------------------------------------------------------------------------
    'Löschen aus der Datenbank
LöschenAusDerDatenbank:
    On Error Resume Next
    If gblnSchreibgeschützt = False Then
        rst1.Delete
        If Err.Number <> 0 Then                     'Gerbing 10.02.2007
            Msg = "Error number=" & Err.Number & vbNewLine
            Msg = Msg & "Error text=" & Err.Description & vbNewLine
            If Err.Number = 3218 Then               'Datensatz ist momentan gesperrt
                'msg = msg & "Wiederholen Sie den Löschversuch zu einem späteren Zeitpunkt"
                Msg = Msg & LoadResString(2326 + Sprache)
            End If
            MsgBox Msg
        End If
    End If
    On Error GoTo 0
End Sub

Private Function ErmittleDateiname(Dateiname As String) As String
    Dim pos As Long
    Dim start As Long
    Dim MeinDateiname As String

    'Der Dateiname wird ermittelt durch Suchen ab rechtem Rand bis zum Punkt
    start = Len(Dateiname) - 2
    Do
        pos = InStr(start, Dateiname, ".")
        If pos <> 0 Then
            MeinDateiname = Mid(Dateiname, 1, pos - 1)
            Exit Do
        End If
        start = start - 1
    Loop
    ErmittleDateiname = MeinDateiname
End Function

Public Sub ZeigeGEOPosition()
    Dim tempDateiname As String
    Dim DateinamenErweiterung As String
    Dim url As String
    Dim Msg As String                                                                   'Gerbing 02.10.2019
    Dim pos As Long
    Dim pos1 As Long
    Dim pos2 As Long
    Dim rc As Long
    Dim antwort As Long                                                                 'Gerbing 02.10.2019
    Dim result As Integer
    Const SW_SHOWNORMAL = 1                                                             'Gerbing 16.10.2018
    Const SW_SHOWMAXIMIZED = 3                                                          'Gerbing 16.10.2018
    
    '7 = Zeige GEO-Position
    'das geht nur für JPG files, in anderen ist keine GEO-Positionen vorhanden
    'ab 29.11.2018 auch für Smartphone MP4-Videos                                       'Gerbing 29.11.2018
    'ab 02.10.2019 für sämtliche files
    'In fotos.exe kann in der Professional Version bei Drücken von Strg+G bei einem mp4 video eine Landkarte gezeigt werden
    'ab 02.10.2019 für sämtliche files
    'jetzt ist nicht mehr der EXIF-Abschnitt die Quelle der Geo-Position sondern die Felder GPSLatitude und GPSLongitude
    'im Gegensatz zu Diashow.exe, dort bleibt der EXIF-Abschnitt die Quelle der Geo-Position
    'wenn mit Fremd-Software (zB Picasa oder GeoSetter) die GPS-Position im EXIF-Abschnitt eingefügt wurde, kann ich diese Werte
    'mit Menü Datei.. -> Feldaktualisierung durch Import-Wiederholung in die Felder GPSLatitude und GPSLongitude übertragen
    
    'es ist eine beliebige Datei
    If gblnProversion Then
        gstrLat = ""
        gstrLong = ""
        On Error Resume Next
        gstrLat = frmGridAndThumb.rsDataGrid.Fields("GPSLatitude")
        gstrLong = frmGridAndThumb.rsDataGrid.Fields("GPSLongitude")
        On Error GoTo 0
        gstrLat = Replace(gstrLat, ",", ".")                                        'Komma in Punkt verwandeln
        gstrLong = Replace(gstrLong, ",", ".")                                      'Komma in Punkt verwandeln
        If gstrLat = "" Or gstrLong = "" Then
            'Wenn keine Geo-Position vorhanden ist kann der Nutzer ein aus OpenStreetMap auswählen
            Msg = LoadResString(3155 + Sprache) & vbNewLine 'keine GEO-Positionen vorhanden
            Msg = Msg & LoadResString(3194 + Sprache) 'Wollen Sie eine Geo-Position eintragen? 'Gerbing 02.10.2019
            antwort = MsgBox(Msg, vbDefaultButton2 + vbYesNo)
            If antwort = vbNo Then
                Exit Sub
            Else
                'Kontrollieren, ob es die Felder GPSLatitude und GPSLongitude in der Tabelle fotos gibt 'Gerbing 05.09.2016
                'wenn nicht, MsgBox zeigen und Abbrechen
                rc = Form1.GPSFelderPrüfen                                                                  'Gerbing 02.10.2019
                If rc = 0 Then Exit Sub                                                                     'Gerbing 02.10.2019
                frmGPSInDatenbankEintragen.Show 1                                                           'Gerbing 02.10.2019
                frmGridAndThumb.btnRefresh_Click                                                            'Gerbing 20.10.2019
                Exit Sub
            End If
        End If
    End If
'   ------------------------------------------------------------------------------------------------------
    'Hierher, wenn eine Geo-Position vorhanden ist
    If gstrLat <> "" And gstrLat <> "0" And gstrLong <> "" And gstrLong <> "0" Then
        'frmGEOPosition.Show 1
        frmStrgG.Show 1                                                             'Gerbing 16.10.2018
        Select Case glngStrgG                                                       'Gerbing 15.04.2020
            Case 1
                frmMap.Show 1
            Case 2
                url = "http://www.openstreetmap.org/?mlat=" & gstrLat & "&mlon=" & gstrLong & "&zoom=16&layers=M?force=tt&hl=de-AT" 'Gerbing 16.10.2018
                ' "Execute" the URL to make the default browser display it.         'Gerbing 16.10.2018
                ShellExecute ByVal 0&, "open", url, _
                    vbNullString, vbNullString, SW_SHOWNORMAL                       'Gerbing 16.10.2018
            Case 3                                                                  'Gerbing 15.04.2020
                'url = "https://maps.google.com/maps?q=50.8359%2C12.9229"
                url = "https://maps.google.com/maps?q=" & gstrLat & "%2C" & gstrLong
                ' "Execute" the URL to make the default browser display it.
                ShellExecute ByVal 0&, "open", url, _
                    vbNullString, vbNullString, SW_SHOWNORMAL
        End Select                                                                  'Gerbing 15.04.2020
    End If
End Sub

Private Function GEOKoordinatenUmrechnen()
    'rc = 0 ohne Fehler
    'rc = 1 Fehler
    
    GEOKoordinatenUmrechnen = 0
    'gstrGEOPosition zusammensetzen                                                 'Gerbing 02.09.2016
    'zB 50.83266,12.45735
    gstrGEOPosition = ""
    gstrLat = ""                                                                    'Gerbing 29.09.2018
    gstrLong = ""                                                                   'Gerbing 29.09.2018
    If GPSLatitudeRef <> "N" Then
        gstrGEOPosition = "-"                                                       '- auf der Südhalbkugel
    End If
    GPSLatitude = Replace(GPSLatitude, ",", ".")                                    'Komma in Punkt verwandeln
    gstrLat = gstrGEOPosition & GPSLatitude                                         'Gerbing 29.09.2018
    gstrGEOPosition = gstrGEOPosition & GPSLatitude & ","
    If GPSLongitudeRef <> "E" Then
        gstrGEOPosition = gstrGEOPosition & "-"                                     '- westlich von Greenwich
        gstrLong = "-"                                                              'Gerbing 29.09.2018
    End If
    GPSLongitude = Replace(GPSLongitude, ",", ".")                                  'Komma in Punkt verwandeln
    gstrGEOPosition = gstrGEOPosition & GPSLongitude
    gstrLong = gstrLong & GPSLongitude                                              'Gerbing 29.09.2018
End Function

Public Function GEOKoordinatenUmrechnenXMP()                                       'Gerbing 08.04.2019
    'zB gstrLatXMP 50,38.7309456N -> 50.64551575
    'zB gstrLongXMP 11,53.9826786E -> 11.89971130
    'Das ist nötig damit OpenStreetMap die GEO-Positionen von verstehen kann
    Dim Grad As String
    Dim Minuten As String                                                   'Gerbing 04.07.2019
    Dim MinutenDouble As Double                                             'Gerbing 04.07.2019
    Dim ESWN As String                                                      'East South West Nord
    Dim Ergebnis As String
    Dim pos As Integer
    Dim locale_id As Long                                                   'Gerbing 04.07.2019
    
    GEOKoordinatenUmrechnenXMP = 0
    pos = InStr(1, gstrLatXMP, ",")                                         'das "," kommt in deutscher und englischer Systemsprache
    Grad = Mid(gstrLatXMP, 1, pos - 1)
    Minuten = Mid(gstrLatXMP, pos + 1, Len(gstrLatXMP) - pos - 1)
    'Wenn Komma als Dezimaltrennzeichen verwendet wird, muss der Punkt im String Minuten in Komma verwandelt werden
    'sonst kommt bei MinutenDouble / 60 Ergebnis=0
    If LocaleInfo(locale_id, LOCALE_SDECIMAL) = "," Then
        Minuten = Replace(Minuten, ".", ",")
    End If
    ESWN = Mid(gstrLatXMP, Len(gstrLatXMP), 1)
    MinutenDouble = CDbl(Minuten)
    MinutenDouble = MinutenDouble / 60
    Ergebnis = Grad + MinutenDouble
    Ergebnis = Replace(Ergebnis, ",", ".")                                  'Gerbing 04.07.2019
    gstrLat = ""
    If ESWN <> "N" Then
        gstrLat = "-"                                                       '- auf der Südhalbkugel
    End If
    gstrLat = gstrLat & Ergebnis                                            'Gerbing 04.07.2019
    '---------------------------
    pos = InStr(1, gstrLongXMP, ",")
    Grad = Mid(gstrLongXMP, 1, pos - 1)
    Minuten = Mid(gstrLongXMP, pos + 1, Len(gstrLongXMP) - pos - 1)
    'Wenn Komma als Dezimaltrennzeichen verwendet wird, muss der Punkt im String Minuten in Komma verwandelt werden
    'sonst kommt bei MinutenDouble / 60 Ergebnis=0
    If LocaleInfo(locale_id, LOCALE_SDECIMAL) = "," Then
        Minuten = Replace(Minuten, ".", ",")
    End If
    ESWN = Mid(gstrLongXMP, Len(gstrLongXMP), 1)
    MinutenDouble = CDbl(Minuten)
    MinutenDouble = MinutenDouble / 60
    Ergebnis = Grad + MinutenDouble
    Ergebnis = Replace(Ergebnis, ",", ".")                                  'Gerbing 04.07.2019
    gstrLong = ""
    If ESWN <> "E" Then
        gstrLong = "-"                                                       '- westlich von Greenwich
    End If
    gstrLong = gstrLong & Ergebnis
End Function

' Bild aus RES-Datei laden und
' als StdPicture-Objekt zurückgeben
    Public Function LoadImageFromRES(ByVal nResID As Long, sType) As StdPicture
    Dim IID_IPicture(3) As Long
    Dim oPicture As IPicture
    Dim bImg() As Byte
    Dim nResult As Long
    Dim oStream As IUnknown
    Dim hGlobal As Long
 
    ' Bild als ByteArray aus RES-Datei laden
    bImg = LoadResData(nResID, sType)
 
    ' Array füllen um den KlassenID (CLSID) IID_IPICTURE
    ' zu simulieren
    IID_IPicture(0) = &H7BF80980
    IID_IPicture(1) = &H101ABF32
    IID_IPicture(2) = &HAA00BB8B
    IID_IPicture(3) = &HAB0C3000
 
     ' Stream erstellen
     Call CreateStreamOnHGlobal(VarPtr(bImg(LBound(bImg))), 0, oStream)
    
     ' OLE IPicture-Objekt erstellen
     nResult = OleLoadPicture(oStream, 0, 0, IID_IPicture(0), oPicture)
    If nResult = 0 Then
        Set LoadImageFromRES = oPicture
    End If
End Function

' Return a piece of locale information.
Private Function LocaleInfo(ByVal locale As Long, ByVal lc_type As Long) As String
Dim Length As Long
Dim buf As String * 1024

    Length = GetLocaleInfo(locale, lc_type, buf, Len(buf))
    LocaleInfo = Left$(buf, Length - 1)
End Function

Public Function GPSFelderPrüfen()                                              'Gerbing 02.10.2019
    GPSFelderPrüfen = 0                                                         'rc=0 bei Fehler
    If Not (gblnVollversion = True And gblnProversion = True) Then              'Gerbing 27.09.2016
        Msg = LoadResString(2335 + Sprache) 'Für diese Funktion benötigen Sie die Professional Version.
        MsgBox Msg
        Exit Function
    End If
    'Kontrollieren, ob es die Felder GPSLatitude und GPSLongitude in der Tabelle fotos gibt 'Gerbing 05.09.2016
    'wenn nicht, MsgBox zeigen und Abbrechen
    SQL = "SELECT top 1 GPSLatitude AS Latitude FROM Fotos;"                'Diese Form versteht sowohl Access als auch SQL-Server
    On Error Resume Next
    rstsql.Close
    Err.Number = 0
    With rstsql
        .Source = SQL
        .ActiveConnection = DBado                                               'Gerbing 23.11.2017
        .CursorType = adOpenForwardOnly
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    If Err.Number <> 0 Then
        'MsgBox "Für die Feld-Aktualisierung durch Import-Wiederholung muss die Datenbank die nutzerdefinierten Felder GPSLatitude und GPSLongitude enthalten"
        MsgBox LoadResString(2348 + Sprache)
        Exit Function
    End If
    SQL = "SELECT top 1 GPSLongitude AS Longitude FROM Fotos;"                'Diese Form versteht sowohl Access als auch SQL-Server
    On Error Resume Next
    rstsql.Close
    Err.Number = 0
    With rstsql
        .Source = SQL
        .ActiveConnection = DBado                                               'Gerbing 23.11.2017
        .CursorType = adOpenForwardOnly
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    If Err.Number <> 0 Then
        rstsql.Close
        'MsgBox "Für die Suche mit GEO-Daten muss die Datenbank die nutzerdefinierten Felder GPSLatitude und GPSLongitude enthalten"
        MsgBox LoadResString(2342 + Sprache)
        Exit Function
    End If
    rstsql.Close
    On Error GoTo 0
    GPSFelderPrüfen = 1
End Function
