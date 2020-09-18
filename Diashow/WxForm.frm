VERSION 5.00
Begin VB.Form WertxForm 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Einstellungen"
   ClientHeight    =   7404
   ClientLeft      =   5256
   ClientTop       =   972
   ClientWidth     =   8196
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7404
   ScaleWidth      =   8196
   Begin VB.Frame FrmSchriftgroesse 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Schriftgröße"
      Height          =   2652
      Left            =   4800
      TabIndex        =   17
      Top             =   4080
      Width           =   3252
      Begin VB.OptionButton OptGross 
         BackColor       =   &H00C0C0C0&
         Caption         =   "groß"
         Height          =   372
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   2412
      End
      Begin VB.OptionButton OptMittel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "mittel"
         Height          =   372
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   2532
      End
      Begin VB.OptionButton OptKlein 
         BackColor       =   &H00C0C0C0&
         Caption         =   "klein"
         Height          =   372
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   2412
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Beim Bildladen vergrößern auf Vollbild"
      Height          =   2652
      Left            =   120
      TabIndex        =   10
      Top             =   4080
      Width           =   4452
      Begin VB.OptionButton optImmer 
         BackColor       =   &H00C0C0C0&
         Caption         =   "immer"
         Height          =   252
         Left            =   240
         TabIndex        =   16
         Top             =   2280
         Value           =   -1  'True
         Width           =   2292
      End
      Begin VB.OptionButton OptKeine 
         BackColor       =   &H00C0C0C0&
         Caption         =   "keine"
         Height          =   372
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   3732
      End
      Begin VB.OptionButton Opt640 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ab Breite 640 Pixel"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   3732
      End
      Begin VB.OptionButton Opt800 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ab Breite 800 Pixel"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   3972
      End
      Begin VB.OptionButton Opt1024 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ab Breite 1024 Pixel"
         Height          =   252
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   3852
      End
      Begin VB.OptionButton Opt1024Oder768 
         BackColor       =   &H00C0C0C0&
         Caption         =   "when image width over 1024 or image hight over 768 pixel"
         Height          =   612
         Left            =   240
         TabIndex        =   11
         Top             =   1680
         Width           =   3732
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Automatik-Wert x (siehe F6)"
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7932
      Begin VB.CommandButton btnNichtStarten 
         Caption         =   "Automatik &nicht starten=F7"
         Height          =   492
         Left            =   4440
         TabIndex        =   7
         Top             =   1920
         Width           =   3372
      End
      Begin VB.CommandButton btnStarten 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Automatik &starten"
         Default         =   -1  'True
         Height          =   492
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   3372
      End
      Begin VB.TextBox txtInterval 
         Height          =   285
         Left            =   2520
         TabIndex        =   5
         Text            =   "3"
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Automatik-Intervall:"
         Height          =   252
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   2292
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Geben Sie einen Wert zwischen 1 und 60 ein. Das ist das Intervall in Sekunden bis das nächste Bild gezeigt wird."
         Height          =   1092
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   7692
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Bildreihenfolge während der automatischen Anzeige"
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   7932
      Begin VB.OptionButton OptAlphabetisch 
         BackColor       =   &H00C0C0C0&
         Caption         =   "alphabetische Reihenfolge der Dateinamen"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   5892
      End
      Begin VB.OptionButton OptZufällig 
         BackColor       =   &H00C0C0C0&
         Caption         =   "zufällige Reihenfolge"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   5892
      End
   End
   Begin VB.CheckBox CheckSpeichernBildPosition 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Beim Bildwechsel individuell vergrößerten/verkleinerten Bildausschnitt beibehalten"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   6840
      Width           =   8052
   End
End
Attribute VB_Name = "WertxForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Function WertxPrüfen(Wertx)
    If Wertx = "" Then
        WertxPrüfen = 1             '1=Fehler
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

    rc = WertxPrüfen(txtInterval.Text)
    If rc <> 0 Then
'        MsgBox "Geben Sie einen Wert zwischen 1 und 60 ein"
        MsgBox LoadResString(2186 + Sprache)
        Exit Sub
    End If
    DiashowForm.Timer1.Interval = txtInterval.Text * 1000
    KeyCode = vbKeyF7
    DiashowForm.F6TimerGestartet = False
    Shift = 0
    Call DiashowForm.Form_KeyDown(KeyCode, Shift)
    Me.Hide
End Sub

Private Sub btnStarten_Click()
    Dim KeyCode As Integer
    Dim Shift As Integer
    Dim rc As Integer
    
    rc = WertxPrüfen(txtInterval.Text)
    If rc <> 0 Then
'        MsgBox "Geben Sie einen Wert zwischen 1 und 60 ein"
        MsgBox LoadResString(2186 + Sprache)
        Exit Sub
    End If
    DiashowForm.Timer1.Interval = txtInterval.Text * 1000
    KeyCode = vbKeyF6
    DiashowForm.F6TimerGestartet = True
    Shift = 0
    Call DiashowForm.Form_KeyDown(KeyCode, Shift)
    Me.Hide
End Sub

Private Sub CheckSpeichernBildPosition_Click()
    Dim i As Long
    
    If CheckSpeichernBildPosition.Value = 0 Then        'Gerbing 29.07.2006
        If DiashowForm.List1U.ListItems.Count > 0 Then
            For i = 0 To DiashowForm.List1U.ListItems.Count - 1
               BildPosList(i).Top = 0
               BildPosList(i).Left = 0
               BildPosList(i).ZoomPercent = 0
            Next i
        End If
    End If
End Sub

Private Sub Form_Load()
    Call AnpassenNutzerWunsch(Me)                               'Gerbing 11.03.2017
    
    Me.Caption = LoadResString(1218 + Sprache)              'Einstellungen
    Label1.Caption = LoadResString(1221 + Sprache)          'Geben Sie einen Wert zwischen 1 und 60 ein. Das ist das Intervall in Sekunden bis das nächste Bild gezeigt wird.
    Label2.Caption = LoadResString(1098 + Sprache)          '&Automatik-Intervall:
    Frame1.Caption = LoadResString(3047 + Sprache)          'Automatik-Wert x (siehe F6)
    btnStarten.Caption = LoadResString(3048 + Sprache)      'Automatik &starten
    btnNichtStarten.Caption = LoadResString(3049 + Sprache) 'Automatik &nicht starten=F7
    Frame3.Caption = LoadResString(3104 + Sprache)          'Bildreihenfolge während der automatischen Anzeige
    Frame5.Caption = LoadResString(3073 + Sprache)          'Beim Bildladen vergrößern auf Vollbild                 'Gerbing 16.03.2009
    FrmSchriftgroesse.Caption = LoadResString(2536 + Sprache)   'Schriftgröße                                       'Gerbing 11.03.2017
    optImmer.Caption = LoadResString(1841 + Sprache)        'immer                                                  'Gerbing 29.04.2013
    OptKeine.Caption = LoadResString(3074 + Sprache)        'Keine
    Opt640.Caption = LoadResString(3075 + Sprache)          'ab Breite 640 Pixel
    Opt800.Caption = LoadResString(3076 + Sprache)          'ab Breite 800 Pixel
    Opt1024.Caption = LoadResString(3077 + Sprache)         'ab Breite 1024 Pixel
    Opt1024Oder768.Caption = LoadResString(3084 + Sprache)  'ab Bildbreite 1024 Pixel oder Bildhöhe 768 Pixel       'Gerbing 16.03.2009
    OptAlphabetisch.Caption = LoadResString(3105 + Sprache) 'alphabetische Reihenfolge der Dateinamen
    OptZufällig.Caption = LoadResString(3106 + Sprache)     'zufällige Reihenfolge
    OptKlein.Caption = LoadResString(1845 + Sprache)        'klein                                                  'Gerbing 11.03.2017
    OptMittel.Caption = LoadResString(1846 + Sprache)       'mittel                                                 'Gerbing 11.03.2017
    OptGross.Caption = LoadResString(1847 + Sprache)        'gross                                                  'Gerbing 11.03.2017
    CheckSpeichernBildPosition.Caption = LoadResString(1222 + Sprache)  'Beim Bildwechsel individuell vergrößerten/verkleinerten Bildausschnitt beibehalten
    
    If PublicCheckForDPI = "1" Then                                                                                 'Gerbing 11.03.2017
        OptKlein.Value = True
    End If
    If PublicCheckForDPI = "2" Then
        OptMittel.Value = True
    End If
    If PublicCheckForDPI = "3" Then
        OptGross.Value = True
    End If
    
    Select Case PublicZoomToFullscreen
        Case 0
            OptKeine = True
        Case 1                                                                              'Gerbing 20.04.2013
            optImmer.Value = True
        Case 640
            Opt640 = True
        Case 800
            Opt800 = True
        Case 1024
            Opt1024 = True
        Case 1024768                                                                        'Gerbing 16.09.2009
            Opt1024Oder768 = True
    End Select
End Sub
        
Private Sub OptGross_Click()                                            'Gerbing 11.03.2017
    Dim frm As Form
    
    PublicCheckForDPI = "3"
    WriteDPI ("3")
    For Each frm In Forms                                               'Gerbing 18.01.2018
        Call AnpassenNutzerWunsch(frm)
    Next
End Sub

Private Sub OptKeine_Click()
    WriteAZF ("0")                                                          'Gerbing 22.08.2007
    PublicZoomToFullscreen = "0"                                            'Gerbing 30.05.2012
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

Private Sub optImmer_Click()
    WriteAZF ("1")                                                        'Gerbing 29.04.2013
    PublicZoomToFullscreen = "1"                                          'Gerbing 29.04.2013
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

Private Sub Form_Paint()
    'SendKeys "%A"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Cancel = True                                                                   'Gerbing 23.10.2007
'    Me.Hide
End Sub
