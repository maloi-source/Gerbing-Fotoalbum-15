VERSION 5.00
Begin VB.Form DuplikatWarnung 
   BackColor       =   &H000000FF&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Duplikat-Warnung"
   ClientHeight    =   2688
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   8328
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2688
   ScaleWidth      =   8328
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton btnKeineWeitereWarnungen 
      Caption         =   "&keine weiteren Warnungen anzeigen"
      Height          =   612
      Left            =   4560
      TabIndex        =   4
      Top             =   1920
      Width           =   3612
   End
   Begin VB.CommandButton btnWeitereWarnungen 
      Caption         =   "&weitere Warnungen anzeigen"
      Default         =   -1  'True
      Height          =   612
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   3612
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Caption         =   "ist ein Duplikat. Duplikate werden nicht übernommen."
      Height          =   492
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   8052
   End
   Begin VB.Label lblDN 
      BorderStyle     =   1  'Fest Einfach
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   8052
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Der Dateiname"
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4812
   End
End
Attribute VB_Name = "DuplikatWarnung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnKeineWeitereWarnungen_Click()
    NeueDatensätzeGenerieren.NichtNochmalWarnen = True
    Unload Me
End Sub

Private Sub btnWeitereWarnungen_Click()
    NeueDatensätzeGenerieren.NichtNochmalWarnen = False
    Unload Me
End Sub

Private Sub Form_Load()
    Call AnpassenNutzerWunsch(Me)                                       'Gerbing 11.03.2017
    Me.Caption = LoadResString(1306 + Sprache)        'Duplikat-Warnung
    Label1.Caption = LoadResString(1307 + Sprache)      'Der Dateiname
    Label3.Caption = LoadResString(1308 + Sprache)      'ist ein Duplikat. Duplikate werden nicht übernommen.
    btnWeitereWarnungen.Caption = LoadResString(1309 + Sprache) '&weitere Warnungen anzeigen
    btnKeineWeitereWarnungen.Caption = LoadResString(1310 + Sprache)    '&keine weiteren Warnungen anzeigen
    
    lblDN.Caption = NeueDatensätzeGenerieren.AktuellerDateiname
End Sub
