VERSION 5.00
Begin VB.Form frmDurchsuchenUnterOrdner 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "GERBING-Diashow"
   ClientHeight    =   2172
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8916
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2172
   ScaleWidth      =   8916
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnAbbrechen 
      Caption         =   "&Abbrechen"
      Height          =   372
      Left            =   6720
      TabIndex        =   6
      Top             =   1680
      Width           =   2052
   End
   Begin VB.CommandButton btnMitUnterOrdnern 
      Caption         =   "mit Unterordnern"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton btnOhneUnterordner 
      Caption         =   "without subfolders"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   """Mit F3 vorwärts- mit F2 rückwärtsblättern"""
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   8772
   End
   Begin VB.Label Label3 
      Caption         =   """Mit F5 können Sie die Liste aller gefundenen Bilder anzeigen"""
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   8772
   End
   Begin VB.Label Label2 
      Caption         =   """durchsucht das Programm den gesamten Ordnern nach weiteren Bildern"""
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   8772
   End
   Begin VB.Label Label1 
      Caption         =   """Wenn Sie ein einzelnes Bild auswählen,"""
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8772
   End
End
Attribute VB_Name = "frmDurchsuchenUnterOrdner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAbbrechen_Click()
    End                                                                     'Gerbing 04.05.2010
End Sub

Private Sub btnMitUnterOrdnern_Click()
    gblnSubdirectories = True                                               'Gerbing 04.03.2013
    btnOhneUnterordner.Enabled = False
    btnMitUnterOrdnern.Enabled = False
    Call DiashowForm.FüllelstFilesImGleichenOrdner
    
    Unload Me
End Sub

Private Sub btnOhneUnterordner_Click()
    gblnSubdirectories = False                                              'Gerbing 04.03.2013
    btnOhneUnterordner.Enabled = False
    btnMitUnterOrdnern.Enabled = False
    Call DiashowForm.FüllelstFilesImGleichenOrdner
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    Call AnpassenNutzerWunsch(Me)                               'Gerbing 11.03.2017

'                Label1.Caption = "Wenn Sie ein einzelnes Bild auswählen,"
'                Label2.Caption = "durchsucht das Programm den gesamten Ordnern nach weiteren Bildern"
'                Label3.Caption = "Mit F5 können Sie die Liste aller gefundenen Bilder anzeigen"
'                Label4.Caption = "Mit F3 vorwärts- mit F2 rückwärtsblättern"
                Label1.Caption = LoadResString(1234 + Sprache)
                Label2.Caption = LoadResString(1231 + Sprache)
                Label3.Caption = LoadResString(1232 + Sprache)
                Label4.Caption = LoadResString(1233 + Sprache)
                btnOhneUnterordner.Caption = LoadResString(3147 + Sprache)
                btnMitUnterOrdnern.Caption = LoadResString(3148 + Sprache)
                btnAbbrechen.Caption = LoadResString(3013 + Sprache)                                    'Gerbing 04.05.2010
                Me.Top = Screen.Height \ 2 - Me.Height \ 2
                Me.Left = Screen.Width \ 2 - Me.Width \ 2
                gblnMessageAusgeben = True                                                               'Gerbing 24.01.2009
End Sub


