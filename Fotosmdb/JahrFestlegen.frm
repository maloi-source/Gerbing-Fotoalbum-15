VERSION 5.00
Begin VB.Form JahrFestlegen 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Jahr festlegen"
   ClientHeight    =   8544
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   12708
   Icon            =   "JahrFestlegen.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8544
   ScaleWidth      =   12708
   StartUpPosition =   1  'Fenstermitte
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   375
      Left            =   120
      ScaleHeight     =   324
      ScaleWidth      =   324
      TabIndex        =   10
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton btnAbbrechen 
      Caption         =   "&Abbrechen"
      Default         =   -1  'True
      Height          =   492
      Left            =   10560
      TabIndex        =   9
      Top             =   7920
      Width           =   2052
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Height          =   492
      Left            =   120
      TabIndex        =   8
      Top             =   7920
      Width           =   2052
   End
   Begin VB.Frame Frame1 
      Caption         =   "Gültigkeit der eingegebenen Jahreszahl"
      Height          =   1572
      Left            =   120
      TabIndex        =   5
      Top             =   6120
      Width           =   12492
      Begin VB.OptionButton OptJedesmalNeu 
         Caption         =   "Jahreszahl jedesmal neu festlegen"
         Height          =   492
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   9012
      End
      Begin VB.OptionButton OptGiltImmer 
         Caption         =   "Diese Jahreszahl gilt für alle weiteren nicht erkannten Jahreszahlen"
         Height          =   492
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   9012
      End
   End
   Begin VB.TextBox txtJahr 
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   5520
      Width           =   732
   End
   Begin VB.Image Image1 
      Height          =   372
      Left            =   600
      Top             =   1800
      Width           =   372
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Geben Sie eine Jahreszahl ein (4-stellig):"
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   5040
      Width           =   9252
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "In dem keine Jahreszahl erkannt werden konnte."
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   9252
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Es wurde folgender Dateiname gefunden"
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9252
   End
   Begin VB.Label lblDN 
      BorderStyle     =   1  'Fest Einfach
      Height          =   612
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   12492
   End
End
Attribute VB_Name = "JahrFestlegen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAbbrechen_Click()
    NeueDatensätzeGenerieren.JahrFestLegenAbgebrochen = True
    Screen.MousePointer = vbHourglass
    Unload Me
End Sub

Private Sub btnOK_Click()
    If Not IsNumeric(txtJahr) Or Len(txtJahr) <> 4 Then
        'MsgBox "Die Jahreszahl muss 4-stellig sein"
        MsgBox LoadResString(1492 + Sprache)
        txtJahr = ""
        txtJahr.SetFocus
        Exit Sub
    End If
    If OptGiltImmer = True Then
        NeueDatensätzeGenerieren.GiltFürAlleFälle = True
    Else
        NeueDatensätzeGenerieren.GiltFürAlleFälle = False
    End If
    NeueDatensätzeGenerieren.JahrFestLegenAbgebrochen = False
    NeueDatensätzeGenerieren.NutzerJahresZahl = txtJahr
    Screen.MousePointer = vbHourglass
    Unload Me
End Sub

Private Sub Form_Load()
    Dim DateinamenErweiterung As String
    
    Call AnpassenNutzerWunsch(Me)                                       'Gerbing 11.03.2017
    Me.Caption = LoadResString(1331 + Sprache)      'Jahr festlegen
    Label1.Caption = LoadResString(1330 + Sprache)  'Es wurde folgender Dateiname gefunden
    Label2.Caption = LoadResString(1332 + Sprache)      'In dem keine Jahreszahl erkannt werden konnte.
    Label3.Caption = LoadResString(1333 + Sprache)      'Geben Sie eine Jahreszahl ein (4-stellig):
    Frame1.Caption = LoadResString(1334 + Sprache)      'Gültigkeit der eingegebenen Jahreszahl
    OptGiltImmer.Caption = LoadResString(1335 + Sprache)        'Diese Jahreszahl gilt für alle weiteren nicht erkannten Jahreszahlen
    OptJedesmalNeu.Caption = LoadResString(1336 + Sprache)      'Jahreszahl jedesmal neu festlegen
    btnOK.Caption = LoadResString(3001 + Sprache)               '&OK
    btnAbbrechen.Caption = LoadResString(3013 + Sprache)        '&Abbrechen
    
    Screen.MousePointer = vbDefault
    lblDN.Caption = NeueDatensätzeGenerieren.AktuellerDateiname
    DateinamenErweiterung = UCase(Right(NeueDatensätzeGenerieren.AktuellerDateiname, 3))
    Select Case DateinamenErweiterung
        Case "BMP", "CUR", "DIB", "EMF", "GIF", "ICO", "JPG", "WMF"         'Gerbing 09.03.2005
            'nur wenn es tatsächlich eine Bilddatei ist
            Call BildAnzeigen(NeueDatensätzeGenerieren.AktuellerDateiname)
    End Select
    txtJahr.Text = Year(Now)                                                'Gerbing 22.12.2019
End Sub


Public Sub BildAnzeigen(EchterStandort)
    Dim Bildbreite As Long
    Dim Bildhöhe As Long
    Dim Image1Top As Long
    Dim Image1Left As Long
    Dim BHV As Double   'BreitenHöhenVerhältnis
    Dim strTemp As String

    strTemp = Replace(EchterStandort, "+:\", AppPath & "\")        'Gerbing 11.04.2005
    Bildbreite = 4000
    Bildhöhe = 3000
    Image1Top = Picture1.top
    Image1Left = Picture1.left

    On Error GoTo 0
    ERR = 0
    Screen.MousePointer = vbHourglass
    Image1.Visible = False
    Picture1.Visible = False
    Picture1.AutoSize = True
    On Error Resume Next
    'Picture1.Picture = LoadPicture(strTemp)
    'Call LoadPictureW(strTemp, Me)
    'Set Picture1.Picture = CreateThumbnailFromFile(strTemp, 100)       'Gerbing 30.09.2015
    Set Picture1.Picture = CreateStdPictureFromFile(strTemp)            'Gerbing 02.10.2015
    If ERR.Number <> 0 Then
        Screen.MousePointer = vbDefault
        Call BildFehler(strTemp)
        Exit Sub
    End If
    On Error GoTo 0
    '-----------------------------------------------------------------------------------------------
    'Untersuchung, ob das Bild größer ist als die Bildbreite/höhe und dessen Konsequenzen
    BHV = Picture1.Width / Picture1.Height

    If Picture1.Width > Bildbreite Or Picture1.Height > Bildhöhe Then
    'wenn das Bild größer ist als Bildbreite/Bildhöhe wird es verkleinert
        Image1.Stretch = True
        Image1.Picture = Picture1.Picture
        Select Case BHV
            Case 1.33 To 1.34
                'das Breitenverhältnis ist 4/3 = 1.33
                Image1.top = Image1Top
                Image1.left = Image1Left
                Image1.Width = Bildbreite
                Image1.Height = Bildhöhe
            Case Is < 1.33
                'das Bild ist zu hoch und zu schmal
                Image1.top = Image1Top
                Image1.left = Image1Left
                Image1.Height = Bildhöhe
                Image1.Width = Bildhöhe * BHV
            Case Else
                'das Bild ist zu niedrig und zu breit
                Image1.top = Image1Top
                Image1.left = Image1Left
                Image1.Width = Bildbreite
                Image1.Height = Bildbreite / BHV
        End Select
    Else
    'wenn das Bild nicht größer ist als Bildbreite/Bildhöhe
        'Bild in links oben im Bildbereich anordnen
        Image1.Stretch = True
        Image1.Picture = Picture1.Picture
        Image1.top = Image1Top
        Image1.left = Image1Left
        Image1.Width = Picture1.Width
        Image1.Height = Picture1.Height
    End If
    Image1.Visible = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub BildFehler(EchterStandort)
    Dim Msg As String
    
'    msg = "Bild kann nicht geladen werden" & NL
'    msg = msg & EchterStandort & NL
'    msg = msg & "Prüfen Sie ob diese Datei existiert" & NL
'    msg = msg & "oder ob es sich um einen verbotenen Dateityp handelt."
    Msg = LoadResString(2056 + Sprache) & vbNewLine
    Msg = Msg & EchterStandort & vbNewLine
    Msg = Msg & LoadResString(2301 + Sprache) & vbNewLine
    Msg = Msg & LoadResString(2302 + Sprache)
    MsgBox Msg, vbInformation
End Sub

