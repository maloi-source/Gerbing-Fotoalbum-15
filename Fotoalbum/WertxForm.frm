VERSION 5.00
Begin VB.Form WertxForm 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Werte einstellen"
   ClientHeight    =   7956
   ClientLeft      =   3276
   ClientTop       =   3948
   ClientWidth     =   11904
   Icon            =   "WertxForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7956
   ScaleWidth      =   11904
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Frame FrmSchriftgroesse 
      Caption         =   "Schriftgroesse"
      Height          =   2532
      Left            =   9120
      TabIndex        =   21
      Top             =   3000
      Width           =   2652
      Begin VB.OptionButton OptGross 
         Caption         =   "groß"
         Height          =   372
         Left            =   240
         TabIndex        =   24
         Top             =   1080
         Width           =   1812
      End
      Begin VB.OptionButton OptMittel 
         Caption         =   "mittel"
         Height          =   372
         Left            =   240
         TabIndex        =   23
         Top             =   720
         Width           =   1692
      End
      Begin VB.OptionButton OptKlein 
         Caption         =   "klein"
         Height          =   372
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   1692
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Videos"
      Height          =   1572
      Left            =   120
      TabIndex        =   12
      Top             =   5760
      Width           =   11652
      Begin VB.OptionButton optOtherExternalPlayer 
         Caption         =   "play videos with other external video player"
         Height          =   252
         Left            =   240
         TabIndex        =   19
         Top             =   1080
         Width           =   7212
      End
      Begin VB.OptionButton optWmp 
         Caption         =   "Videos abspielen mit internem Mediaplayer 10"
         Height          =   372
         Left            =   240
         TabIndex        =   14
         ToolTipText     =   "Es wird wmp.dll benutzt"
         Top             =   240
         Width           =   11292
      End
      Begin VB.OptionButton OptMediaplayer 
         Caption         =   "play videos with external windows media player"
         Height          =   432
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   11292
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Beim Bildladen vergrößern auf Vollbild"
      Height          =   2532
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   8892
      Begin VB.OptionButton optImmer 
         Caption         =   "immer"
         Height          =   252
         Left            =   240
         TabIndex        =   20
         Top             =   2160
         Width           =   1812
      End
      Begin VB.OptionButton Opt1024Oder768 
         Caption         =   "when image width over 1024 or image hight over 768 pixel"
         Height          =   372
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   8652
      End
      Begin VB.OptionButton Opt1024 
         Caption         =   "ab Breite 1024 Pixel"
         Height          =   372
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   8652
      End
      Begin VB.OptionButton Opt800 
         Caption         =   "ab Breite 800 Pixel"
         Height          =   372
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   8652
      End
      Begin VB.OptionButton Opt640 
         Caption         =   "ab Breite 640 Pixel"
         Height          =   372
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   8532
      End
      Begin VB.OptionButton OptKeine 
         Caption         =   "keine"
         Height          =   372
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   8412
      End
   End
   Begin VB.CheckBox CheckSpeichernBildPosition 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Beim Bildwechsel individuell vergrößerten/verkleinerten Bildausschnitt beibehalten"
      Height          =   372
      Left            =   120
      TabIndex        =   6
      Top             =   7440
      Width           =   11652
   End
   Begin VB.Frame Frame1 
      Caption         =   "Automatik-Wert x (siehe Strg+F6)"
      Height          =   2652
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11652
      Begin VB.Frame Frame6 
         Height          =   1092
         Left            =   240
         TabIndex        =   16
         Top             =   1440
         Width           =   5772
         Begin VB.OptionButton OptAufsteigend 
            Caption         =   "Automatik mit aufsteigender Reihenfolge"
            Height          =   372
            Left            =   240
            TabIndex        =   18
            Top             =   600
            Value           =   -1  'True
            Width           =   5412
         End
         Begin VB.OptionButton OptZufall 
            Caption         =   "Automatik mit Zufallsreihenfolge"
            Height          =   372
            Left            =   240
            TabIndex        =   17
            Top             =   240
            Width           =   5412
         End
      End
      Begin VB.TextBox txtInterval 
         Height          =   372
         Left            =   3000
         TabIndex        =   3
         Text            =   "3"
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton btnStarten 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Automatik &starten"
         Default         =   -1  'True
         Height          =   375
         Left            =   6480
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Grafisch
         TabIndex        =   2
         Top             =   1560
         Width           =   3612
      End
      Begin VB.CommandButton btnNichtStarten 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Automatik &nicht starten=F7"
         Height          =   375
         Left            =   6480
         Style           =   1  'Grafisch
         TabIndex        =   1
         Top             =   2160
         Width           =   3612
      End
      Begin VB.Label Label1 
         Caption         =   $"WertxForm.frx":038A
         Height          =   732
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   11172
      End
      Begin VB.Label Label2 
         Caption         =   "&Automatik-Intervall:"
         Height          =   372
         Left            =   360
         TabIndex        =   4
         Top             =   1080
         Width           =   2652
      End
   End
End
Attribute VB_Name = "WertxForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim blnIchKommeAusFormLoad As Boolean

Private Function WertxPrüfen(Wertx)
    If Wertx = "" Then
'        WertxPrüfen = 1             '1=Fehler
        txtInterval.Text = 3        'Gerbing 23.06.2011
        Exit Function
    End If
    If IsNumeric(Wertx) <> True Then
        WertxPrüfen = 1             '1=Fehler
        Exit Function
    End If
    If Wertx < 1 Or Wertx > 60 Then
        WertxPrüfen = 1             '1=Fehler
        Exit Function
    End If
    If Len(Wertx) > 2 Then
        WertxPrüfen = 1             '1=Fehler
        Exit Function
    End If
    WertxPrüfen = 0                 '0=kein Fehler
End Function

Private Sub btnNichtStarten_Click()
    Dim KeyCode As Integer
    Dim Shift As Integer
    Dim rc As Integer

    'rc = WertxPrüfen(txtInterval.Text)                         'Gerbing 23.06.2011
    'If rc <> 0 Then
        'MsgBox "Geben Sie einen Wert zwischen 1 und 60 ein"    'Gerbing 08.11.2005
    If txtInterval.Text = "" Then                               'Gerbing 23.06.2011
        MsgBox LoadResString(2186 + Sprache)
        txtInterval.Text = 3
        Exit Sub
    End If
    Form1.lngTimer1Interval = txtInterval.Text * 1000
        KeyCode = vbKeyF7
    Shift = 0
    Unload Me
    Call Form1.Form_KeyDown(KeyCode, Shift)         'Gerbing 29.03.2012
End Sub

Private Sub btnStarten_Click()
    Dim KeyCode As Integer
    Dim Shift As Integer
    Dim rc As Integer
    
    'rc = WertxPrüfen(txtInterval.Text)                         'Gerbing 23.06.2011
    'If rc <> 0 Then
    If txtInterval.Text = "" Then                               'Gerbing 23.06.2011
        'MsgBox "Geben Sie einen Wert zwischen 1 und 60 ein"    'Gerbing 08.11.2005
        MsgBox LoadResString(2186 + Sprache)
        txtInterval.Text = 3
        Exit Sub
    End If
    Form1.lngTimer1Interval = txtInterval.Text * 1000
    KeyCode = vbKeyF6
    Shift = vbCtrlMask      'Gerbing 03.11.2004 Strg+F6 gleichzeitig
    Unload Me
    Call Form1.Form_KeyDown(KeyCode, Shift)         'Gerbing 29.03.2012
End Sub

Private Sub CheckSpeichernBildPosition_Click()
    Dim i As Long
    
    If CheckSpeichernBildPosition.Value = 0 Then        'Gerbing 29.07.2006
        gblnCheckSpeichernBildPosition = False
        For i = 0 To Query.RecordCount - 1
            BildPosList(i).Top = 0
            BildPosList(i).Left = 0
            BildPosList(i).ZoomPercent = 0
            Mid(BildPosList(i).Dateiname, 1, 1) = "?"
        Next i
    Else
        gblnCheckSpeichernBildPosition = True
    End If
End Sub

Private Sub Form_Load()
    blnIchKommeAusFormLoad = True
    Call AnpassenNutzerWunsch(Me)                           'Gerbing 11.03.2017
    gblnMsg = False                                         'Gerbing 02.09.2008
    If Query.chkFensterGrößeÄnderbar.Value = 1 Then         'Gerbing 06.12.2005
        Me.Top = Form1.Top                                  'Gerbing 06.12.2006
        Me.Left = Form1.Left
    Else
        Me.Top = 300                                        'Gerbing 25.10.2007
        Me.Left = 300
    End If

    Me.Caption = LoadResString(1096 + Sprache)              'Werte einstellen
    Label1.Caption = LoadResString(1097 + Sprache)          'Geben Sie einen Wert zwischen 1 und 60 ein. Das ist das Intervall in Sekunden bis das nächste Bild gezeigt wird. Dateien mit Link-Dateityp werden bei der automatischen Anzeige ignoriert.
    Label2.Caption = LoadResString(1098 + Sprache)          '&Automatik-Intervall:
    Frame1.Caption = LoadResString(3047 + Sprache)          'Automatik-Wert x (siehe Strg+F6)  Gerbing 21.08.2007
    Frame3.Caption = LoadResString(3073 + Sprache)          'Beim Bildladen vergrößern auf Vollbild
    FrmSchriftgroesse.Caption = LoadResString(2536 + Sprache)   'Schriftgröße                                       'Gerbing 11.03.2017
    OptKeine.Caption = LoadResString(3074 + Sprache)        'Keine
    Opt640.Caption = LoadResString(3075 + Sprache)          'ab Breite 640 Pixel
    Opt800.Caption = LoadResString(3076 + Sprache)          'ab Breite 800 Pixel
    Opt1024.Caption = LoadResString(3077 + Sprache)         'ab Breite 1024 Pixel
    Opt1024Oder768.Caption = LoadResString(3084 + Sprache)  'ab Bildbreite 1024 Pixel oder Bildhöhe 768 Pixel       'Gerbing 16.03.2009
    optImmer.Caption = LoadResString(1841 + Sprache)        'immer                                                  'Gerbing 29.04.2013
    OptZufall.Caption = LoadResString(3089 + Sprache)       'Automatik mit Zufallsreihenfolge                       'Gerbing 22.11.2010
    OptAufsteigend.Caption = LoadResString(3090 + Sprache)  'Automatik mit aufsteigender Reihenfolge
    gblnWertxFormOptZufall = False                                                                                  'Gerbing 11.08.2012
    OptKlein.Caption = LoadResString(1845 + Sprache)        'klein                                                  'Gerbing 11.03.2017
    OptMittel.Caption = LoadResString(1846 + Sprache)       'mittel                                                 'Gerbing 11.03.2017
    OptGross.Caption = LoadResString(1847 + Sprache)        'gross                                                  'Gerbing 11.03.2017
    optWmp.Caption = LoadResString(3081 + Sprache)          'Videos abspielen mit internem Mediaplayer 7 oder aufwärts   'Gerbing 06.05.2009
    OptMediaplayer.Caption = LoadResString(3079 + Sprache)  'externen Windows Mediaplayer benutzen
    optOtherExternalPlayer.Caption = LoadResString(3151 + Sprache) 'play videos with other external video player    'Gerbing 28.12.2011
    optWmp.tooltipText = LoadResString(3083 + Sprache)      'Es wird wmp.dll benutzt
    
    btnStarten.Caption = LoadResString(3048 + Sprache)      'Automatik &starten
    btnNichtStarten.Caption = LoadResString(3049 + Sprache)    'Automatik &nicht starten=F7
    CheckSpeichernBildPosition.Caption = LoadResString(1222 + Sprache)  'Beim Bildwechsel individuell vergrößerten/verkleinerten Bildausschnitt beibehalten
    'CheckSpeichernBildPosition.ToolTipText = LoadResString(1115 + Sprache)       'Diese Einstellung kann nur gewählt werden, wenn weniger als 100 Ergebnissätze vorhanden sind
    
    Call ReadFotosIniFile                                       'Gerbing 29.03.2012
    
    If PublicAutomaticInterval <> 0 Then                        'Gerbing 28.03.2010
        txtInterval.Text = PublicAutomaticInterval              'Gerbing 22.08.2007
    End If
        
    'optOtherExternalPlayer = True                                                           'Gerbing 28.12.2011
    Select Case PublicPlayVideosWith                                                        'Gerbing 01.09.2008
        Case 10
            optWmp.Value = True
        Case "external"
            OptMediaplayer.Value = True
        Case Else                                                                           'Gerbing 28.12.2011
            optOtherExternalPlayer.Value = True
    End Select
    
    Select Case PublicZoomToFullscreen
        Case 0
            OptKeine.Value = True
        Case 1                                                                              'Gerbing 20.04.2013
            optImmer.Value = True
        Case 640
            Opt640.Value = True
        Case 800
            Opt800.Value = True
        Case 1024
            Opt1024.Value = True
        Case 1024768                                                                        'Gerbing 16.09.2009
            Opt1024Oder768.Value = True
    End Select
    
    If PublicCheckForDPI = "1" Then                                                         'Gerbing 11.03.2017
        OptKlein.Value = True
    End If
    If PublicCheckForDPI = "2" Then
        OptMittel.Value = True
    End If
    If PublicCheckForDPI = "3" Then
        OptGross.Value = True
    End If

    If Query.OKGewählt = False Then
        btnStarten.Visible = False
        btnNichtStarten.Visible = False
    End If
    gblnMsg = True                                                                          'Gerbing 02.09.2008
    If gblnCheckSpeichernBildPosition = True Then
        CheckSpeichernBildPosition.Value = 1
    Else
        CheckSpeichernBildPosition.Value = 0
    End If
    blnIchKommeAusFormLoad = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If gblnComeFromF8 = True Then Exit Sub                                  'Gerbing 31.05.2013
End Sub

Private Sub Opt1024oder768_Click()
    WriteAZF ("1024768")                                                    'Gerbing 16.03.2009
    PublicZoomToFullscreen = "1024768"                                      'Gerbing 30.05.2012
End Sub

Private Sub Opt1024_Click()
    WriteAZF ("1024")                                                       'Gerbing 22.08.2007
    PublicZoomToFullscreen = "1024"                                         'Gerbing 30.05.2012
End Sub

Private Sub Opt640_Click()
    WriteAZF ("640")                                                        'Gerbing 22.08.2007
    PublicZoomToFullscreen = "640"                                          'Gerbing 30.05.2012
End Sub

Private Sub Opt800_Click()
    WriteAZF ("800")                                                        'Gerbing 22.08.2007
    PublicZoomToFullscreen = "800"                                          'Gerbing 30.05.2012
End Sub

Private Sub OptAufsteigend_Click()
    gblnWertxFormOptZufall = False
End Sub

Private Sub optImmer_Click()
    WriteAZF ("1")                                                        'Gerbing 29.04.2013
    PublicZoomToFullscreen = "1"                                          'Gerbing 29.04.2013
End Sub
        
Private Sub OptGross_Click()                                            'Gerbing 11.03.2017
    Dim frm As Form
    
    PublicCheckForDPI = "3"
    WriteDPI ("3")
    For Each frm In Forms                                               'Gerbing 18.01.2018
        Call AnpassenNutzerWunsch(frm)
    Next
End Sub

Private Sub OptKlein_Click()                                            'Gerbing 11.03.2017
    Dim frm As Form
    
    PublicCheckForDPI = "1"
    WriteDPI ("1")
    For Each frm In Forms                                               'Gerbing 18.01.2018
        Call AnpassenNutzerWunsch(frm)
    Next
End Sub

Private Sub OptMittel_Click()                                           'Gerbing 11.03.2017
    Dim frm As Form
    
    PublicCheckForDPI = "2"
    WriteDPI ("2")
    For Each frm In Forms                                               'Gerbing 18.01.2018
        Call AnpassenNutzerWunsch(frm)
    Next
End Sub

Private Sub OptZufall_Click()
    gblnWertxFormOptZufall = True
End Sub

Private Sub OptKeine_Click()
    WriteAZF ("0")                                                          'Gerbing 22.08.2007
    PublicZoomToFullscreen = "0"                                            'Gerbing 30.05.2012
End Sub

Private Sub OptMediaplayer_Click()
    Dim strTemp As String
    Dim Pos As Long
    Dim FileName As String
    
    On Error Resume Next
    'Call Form1.MediaPlayerStop          auskommentiert 23.10.2013                          'Gerbing 01.09.2008
    On Error GoTo 0
    gstrWmplayerFolder = getSpecialFolder(CSIDL_PROGRAM_FILES_COMMONX86)
    'liefert zB C:\Programme\Gemeinsame Dateien
    Pos = InStr(4, gstrWmplayerFolder, "\")
    If Pos <> 0 Then
        gstrWmplayerFolder = Left(gstrWmplayerFolder, Pos) & "Windows Media Player" & "\wmplayer.exe"
'        strTemp = Dir(gstrWmplayerFolder)
'        If strTemp = "" Then Pos = 0
        If file_path_exist(gstrWmplayerFolder) = False Then Pos = 0
    End If
    If Pos = 0 Then
        FileName = ShowOpenUnicodeExternalVideoPlayer(Me)
        FileName = ConvertFileName(FileName)
        Pos = InStr(1, FileName, "wmplayer.exe")
        If Pos = 0 Then GoTo ErrHandler
        gstrWmplayerFolder = FileName
    End If
    'Query.chkFensterGrößeÄnderbar.Value = 1                                 'Gerbing 19.10.2007 06.06.2012
    'Fenstergröße änderbar ja
    Form1.WindowState = 0                      '0=normal
    WriteMPVW ("external")                                                                  'Gerbing 01.09.2008
    PublicPlayVideosWith = "external"
    Exit Sub
ErrHandler:
  'User pressed the Cancel button
  optWmp.Value = True
  Exit Sub
End Sub

Private Sub optOtherExternalPlayer_Click()                                  'Gerbing 28.12.2011
    Dim FileName As String
    
    If blnIchKommeAusFormLoad = True Then                                   'Gerbing 29.03.2012
        Exit Sub
    End If
'    If InStr(1, PublicPlayVideosWith, ".exe", vbTextCompare) = 0 Then      'Gerbing 28.12.2011
'        Exit Sub
'    End If
    On Error Resume Next                                                    'Gerbing 05.08.2013
    'Call Form1.MediaPlayerStop          auskommentiert 23.10.2013                          'Gerbing 01.09.2008
    On Error GoTo 0                                                         'Gerbing 05.08.2013
    FileName = ShowOpenUnicodeOtherVideoPlayer(Me)
    FileName = ConvertFileName(FileName)
    If FileName = "" Then GoTo ErrHandler
    gstrWmplayerFolder = FileName
    'Query.chkFensterGrößeÄnderbar.Value = 1                                'Gerbing 06.06.2012
    'Fenstergröße änderbar ja
    Form1.WindowState = 0                      '0=normal
    WriteMPVW (gstrWmplayerFolder)
    PublicPlayVideosWith = gstrWmplayerFolder
    Exit Sub
ErrHandler:
  'User pressed the Cancel button
  optWmp.Value = True
  Exit Sub
End Sub

Private Sub optWmp_Click()
    'Call Form1.MediaPlayerStop          auskommentiert 23.10.2013                          'Gerbing 01.09.2008
    WriteMPVW ("10")                                                                        'Gerbing 01.09.2008
    PublicPlayVideosWith = "10"                                                             'Gerbing 01.09.2008
End Sub

Private Sub txtInterval_Change()
    Dim rc As Integer

    rc = WertxPrüfen(txtInterval.Text)
    If rc <> 0 Then
        'MsgBox "Geben Sie einen Wert zwischen 1 und 60 ein"    'Gerbing 08.11.2005
        MsgBox LoadResString(2186 + Sprache)
        txtInterval.Text = 3
        Exit Sub
    End If
    WriteAAI (txtInterval.Text)                                 'Gerbing 22.08.2007
    PublicAutomaticInterval = txtInterval.Text
    Form1.lngTimer1Interval = txtInterval.Text * 1000
End Sub
