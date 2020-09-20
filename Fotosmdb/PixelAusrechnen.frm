VERSION 5.00
Begin VB.Form PixelAusrechnen 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Berechnen Pixel"
   ClientHeight    =   5148
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   12648
   Icon            =   "PixelAusrechnen.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5148
   ScaleWidth      =   12648
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton btnReturn 
      Caption         =   "&Prüfen1 abbrechen"
      Height          =   612
      Left            =   9840
      TabIndex        =   6
      Top             =   4440
      Width           =   2772
   End
   Begin VB.CommandButton btnNichtKontrollieren 
      Caption         =   "nicht &kontrollieren"
      Height          =   612
      Left            =   6600
      TabIndex        =   3
      Top             =   4440
      Width           =   2772
   End
   Begin VB.TextBox txtMsg 
      BackColor       =   &H8000000F&
      Height          =   3612
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "PixelAusrechnen.frx":038A
      Top             =   720
      Width           =   12372
   End
   Begin VB.CommandButton btnNurNeueAusrechnen 
      Caption         =   "&Nur neue Dateien kontrollieren"
      Default         =   -1  'True
      Height          =   612
      Left            =   3360
      TabIndex        =   1
      Top             =   4440
      Width           =   2772
   End
   Begin VB.CommandButton btnAlleAusrechnen 
      Caption         =   "&Alle Dateien kontrollieren"
      Height          =   612
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Width           =   2772
   End
   Begin VB.Label lblDatum 
      BorderStyle     =   1  'Fest Einfach
      Height          =   372
      Left            =   5040
      TabIndex        =   5
      Top             =   120
      Width           =   2052
   End
   Begin VB.Label Label1 
      Caption         =   "Datum der letzten Berechnung:"
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4692
   End
End
Attribute VB_Name = "PixelAusrechnen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAlleAusrechnen_Click()
    Form1.blnMitBH = True
    Form1.blnNurNeue = False
    Form1.blnReturn = False                                                                 'Gerbing 22.11.2008
    Unload Me
End Sub

Private Sub btnNichtKontrollieren_Click()
    Form1.blnMitBH = False
    Form1.blnNurNeue = False
    Form1.blnReturn = False                                                                 'Gerbing 22.11.2008
    Unload Me
End Sub

Private Sub btnNurNeueAusrechnen_Click()
    Form1.blnMitBH = True
    Form1.blnNurNeue = True
    Form1.blnReturn = False                                                                 'Gerbing 22.11.2008
    Unload Me
End Sub

Private Sub btnReturn_Click()
    Form1.blnReturn = True                                                                  'Gerbing 22.11.2008
    Unload Me
End Sub

Private Sub Form_Load()
    Dim Msg As String
    Dim NL As String
    
    Call AnpassenNutzerWunsch(Me)                                                           'Gerbing 11.03.2017
    'Me.Caption = "Berechnen Pixel"
    Me.Caption = LoadResString(2219 + Sprache)
    
'    msg = LoadResString(1443 + Sprache)     'Prüfen1                                       'Gerbing 15.12.2008
'    msg = msg & ":"
'    msg = msg & LoadResString(1420 + Sprache)  'Prüfen1 = ob jede im Feld Dateiname eingetragene Foto-Datei  wirklich existiert." & NL
'    Msg = MSG & "Zusätzlich können Sie die Bildbreite und Bildhöhe für alle nativen Bilddateien und für Videos" & NL
'    Msg = Msg & "in die Datenbankfelder BreitePixel und HoehePixel berechnen lassen." & NL
'    Msg = Msg & "Das ist nicht nötig, wenn Sie keine Bildgrößenänderungen gemacht haben." & NL
'    Msg = Msg & "Die Arbeitsdauer von Prüfen1 verlängert sich beim Überprüfen aller Dateien mehrfach" & NL
'    Msg = Msg & "Wenn nur einzelne Dateien neu berechnet werden müssen, die ein aktuelleres Datum haben als der letzte Berechnungsvorgang," & NL
'    Msg = Msg & "verlängert sich die Arbeitsdauer kaum." & NL & NL
    
'    Msg = Msg & "Wie wollen Sie die Datenbankfelder BreitePixel und HoehePixel berechnen lassen?" & NL
    NL = vbNewLine
    
    Msg = LoadResString(1443 + Sprache)     'Prüfen1                                        'Gerbing 15.12.2008
    Msg = Msg & ":"
    Msg = Msg & LoadResString(1420 + Sprache) & NL & NL  'Prüfen1 = ob jede im Feld Dateiname eingetragene Foto-Datei  wirklich existiert."
    Msg = Msg & LoadResString(1388 + Sprache) & NL
    'msg = msg & LoadResString(1389 + Sprache) & NL
    Msg = Msg & LoadResString(1390 + Sprache) & NL
    Msg = Msg & LoadResString(1391 + Sprache) & NL
    Msg = Msg & LoadResString(1392 + Sprache) & NL
    Msg = Msg & LoadResString(1393 + Sprache) & NL
    Msg = Msg & LoadResString(1394 + Sprache) & NL & NL
    
    'Msg = Msg & "Wie wollen Sie die Datenbankfelder BreitePixel und HoehePixel eintragen lassen?" & NL
    Msg = Msg & LoadResString(2209 + Sprache) & NL
    txtMsg = Msg
    
    btnNurNeueAusrechnen.Caption = LoadResString(2215 + Sprache)    '&Nur neue Dateien berechnen
    btnAlleAusrechnen.Caption = LoadResString(2216 + Sprache)           '&Alle Dateien berechnen
    btnNichtKontrollieren.Caption = LoadResString(2217 + Sprache)       'keine &berechnen
    btnReturn.Caption = LoadResString(2278 + Sprache)                   'Prüfen1 abbrechen  Gerbing 22.11.2008
    Label1.Caption = LoadResString(2218 + Sprache)                      'Datum der letzten Berechnung:
    If Not Form1.rstDBH.EOF Then
        lblDatum = Form1.rstDBH.Fields("DatumBreiteHoehe")
    End If
End Sub

Private Sub Form_Paint()
    btnNurNeueAusrechnen.SetFocus
End Sub

