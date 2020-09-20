VERSION 5.00
Begin VB.Form frmFolderIsNothing 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "GERBING Fotoalbum"
   ClientHeight    =   2388
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   5784
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2388
   ScaleWidth      =   5784
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   492
      Left            =   2400
      TabIndex        =   2
      Top             =   1680
      Width           =   1092
   End
   Begin VB.CheckBox chkDieseNachrichtNichtMehrZeigen 
      Caption         =   "Diese Nachricht nicht mehr zeigen"
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   5292
   End
   Begin VB.Label lblHinweisWindows10 
      AutoSize        =   -1  'True
      Caption         =   "Video-Thumbnails können Sie erst ab Windows 10 erzeugen"
      Height          =   552
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   5376
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmFolderIsNothing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnOK_Click()
    Unload Me
End Sub

Private Sub chkDieseNachrichtNichtMehrZeigen_Click()
    If chkDieseNachrichtNichtMehrZeigen.Value = 1 Then
        gblnDieseNachrichtNichtMehrZeigen = True
    End If
End Sub

Private Sub Form_Load()
    Call AnpassenNutzerWunsch(Me)                                   'Gerbing 11.03.2017
    Call AnpassenHeadFont(frmGridAndThumb.DBGridNeu)                'Gerbing 23.06.2011

    lblHinweisWindows10.Caption = LoadResString(1131 + Sprache)
    chkDieseNachrichtNichtMehrZeigen.Caption = LoadResString(1132 + Sprache)
End Sub

