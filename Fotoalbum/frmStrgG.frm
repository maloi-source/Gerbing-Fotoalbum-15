VERSION 5.00
Begin VB.Form frmStrgG 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Geo-Position"
   ClientHeight    =   2892
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   7800
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2892
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   492
      Left            =   3120
      TabIndex        =   3
      Top             =   2280
      Width           =   1332
   End
   Begin VB.Frame Frame1 
      Height          =   2052
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7572
      Begin VB.OptionButton OptGoogleMaps 
         Caption         =   "Google Maps Anzeigen im Browser"
         Height          =   372
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   6732
      End
      Begin VB.OptionButton optBrowser 
         Caption         =   "OSM Anzeigen im Browser(dauert länger)"
         Height          =   372
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   7092
      End
      Begin VB.OptionButton optProgrammFenster 
         Caption         =   "OSM Anzeigen im Programmfenster (geht schneller)"
         Height          =   372
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7332
      End
   End
End
Attribute VB_Name = "frmStrgG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call AnpassenNutzerWunsch(Me)                                               'Gerbing 16.10.2018
    optBrowser.Caption = LoadResString(1133 + Sprache)                          'Anzeigen im Browser(dauert länger)
    optProgrammFenster.Caption = LoadResString(1134 + Sprache)                  'Anzeigen im Programmfenster(geht schneller)
    OptGoogleMaps.Caption = LoadResString(1160 + Sprache)                       'Google Maps Anzeigen im Browser    'Gerbing 15.04.2020
    Me.Caption = LoadResString(3162 + Sprache)                                  'GEO-Position anzeigen
End Sub

Private Sub optBrowser_Click()
    glngStrgG = 2
End Sub

Private Sub OptGoogleMaps_Click()
    glngStrgG = 3                                                               'Gerbing 15.04.2020
End Sub

Private Sub optProgrammFenster_Click()
    glngStrgG = 1
End Sub
