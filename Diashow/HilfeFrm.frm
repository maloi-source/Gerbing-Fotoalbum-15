VERSION 5.00
Begin VB.Form HilfeBoxForm 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Hilfe zu den Bedien-Tasten"
   ClientHeight    =   6384
   ClientLeft      =   3120
   ClientTop       =   3060
   ClientWidth     =   9900
   Icon            =   "HilfeFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6384
   ScaleWidth      =   9900
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblStrgG 
      BackColor       =   &H00C0C0C0&
      Caption         =   """Strg+G - zeige die Geo-Position"""
      Height          =   372
      Left            =   240
      TabIndex        =   16
      Top             =   4440
      Width           =   9612
   End
   Begin VB.Label lblAlt 
      BackColor       =   &H00C0C0C0&
      Caption         =   """Alt - Menüleiste ein/aus"
      Height          =   372
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Width           =   9492
   End
   Begin VB.Label LabelStrgC 
      BackColor       =   &H00C0C0C0&
      Caption         =   """Strg+C - kopiert das Bild"""
      Height          =   372
      Left            =   240
      TabIndex        =   14
      Top             =   4800
      Width           =   9492
   End
   Begin VB.Label LabelStrgO 
      BackColor       =   &H00C0C0C0&
      Caption         =   """Strg+O - schaltet die Rechteck-Lupe wieder aus"""
      Height          =   372
      Left            =   240
      TabIndex        =   13
      Top             =   5520
      Width           =   9456
   End
   Begin VB.Label LabelStrgZ 
      BackColor       =   &H00C0C0C0&
      Caption         =   """Strg+Z - macht die Rechteck-Lupe scharf, vorher ist sie nicht benutzbar"""
      Height          =   372
      Left            =   240
      TabIndex        =   12
      Top             =   5160
      Width           =   9456
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0C0C0&
      Caption         =   """F11 - zeigt Bildgröße und Position des Mauszeigers"""
      Height          =   372
      Left            =   240
      TabIndex        =   11
      Top             =   4080
      Width           =   9456
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      Caption         =   """F8 - Neue Diashow zusammenstellen"""
      Height          =   372
      Left            =   240
      TabIndex        =   10
      Top             =   3000
      Width           =   9456
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0C0&
      Caption         =   """F10 - Dialog-Fenster zum Einstellen von Werten"
      Height          =   372
      Left            =   240
      TabIndex        =   9
      Top             =   3720
      Width           =   9456
   End
   Begin VB.Label lblAltF4 
      BackColor       =   &H00C0C0C0&
      Caption         =   """Alt+F4 - Beenden"""
      Height          =   372
      Left            =   240
      TabIndex        =   8
      Top             =   5880
      Width           =   9492
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
      Caption         =   """F9 - Mauszeiger sichtbar/unsichtbar"""
      Height          =   372
      Left            =   240
      TabIndex        =   7
      Top             =   3360
      Width           =   9456
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      Caption         =   """F7 - Stoppe Automatik"""
      Height          =   372
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   9456
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      Caption         =   """F6 - Starte Automatik neues Bild nach x Sekunden (siehe F10)"""
      Height          =   372
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   9456
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   """F5 - zeigt die Dateinamen, EXIF/IPTC-Felder, Klick wählt ein Bild aus"""
      Height          =   372
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   9456
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   """F4 - Zoom In (Vergrößern)"""
      Height          =   372
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   9456
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   """F3 - Gehe ein Bild vorwärts (oder Taste Alt und ->)"""
      Height          =   372
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   9456
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   """F2 - Gehe ein Bild zurück (oder Taste Alt und <-)"""
      Height          =   372
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   9456
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   """F1 - Zoom Out (Verkleinern)"""
      Height          =   372
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   9456
   End
End
Attribute VB_Name = "HilfeBoxForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim KeyCode As Integer
    Dim Shift As Integer
    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
        ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, _
        ByVal wFlags As Long) As Long

    ' SetWindowPos Flags
    Private Const SWP_NOSIZE = &H1
    Private Const SWP_NOMOVE = &H2
    Private Const SWP_NOZORDER = &H4
    Private Const SWP_NOREDRAW = &H8
    Private Const SWP_NOACTIVATE = &H10
    Private Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
    Private Const SWP_SHOWWINDOW = &H40
    Private Const SWP_HIDEWINDOW = &H80
    Private Const SWP_NOCOPYBITS = &H100
    Private Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
    
    Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
    Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
    
    ' SetWindowPos() hwndInsertAfter values
    Private Const HWND_TOP = 0
    Private Const HWND_BOTTOM = 1
    Private Const HWND_TOPMOST = -1
    Private Const HWND_NOTOPMOST = -2

Private Sub Form_Load()
    Dim i As Long
    
    Call AnpassenNutzerWunsch(Me)                               'Gerbing 11.03.2017
    
    Me.Caption = LoadResString(1203 + Sprache)   'Hilfe zu den Bedien-Tasten
    lblAlt.Caption = LoadResString(3111 + Sprache)  '"Alt - Menü ein/aus"                                   'Gerbing 29.01.2018
    Label1.Caption = LoadResString(1204 + Sprache)  '"F1 - Zoom Out (Verkleinern)"
    Label2.Caption = LoadResString(1205 + Sprache)  '"F2 - Gehe ein Bild zurück (oder Taste Alt und <-)"
    Label3.Caption = LoadResString(1206 + Sprache)  '"F3 - Gehe ein Bild vorwärts (oder Taste Alt und ->)"
    Label4.Caption = LoadResString(1207 + Sprache)  '"F4 - Zoom In (Vergrößern)"
    Label5.Caption = LoadResString(1208 + Sprache)  '"F5 - zeigt die Dateinamen, EXIF/IPTC-Felder, Klick wählt ein Bild aus"
    Label6.Caption = LoadResString(1209 + Sprache)  '"F6 - Starte Automatik neues Bild nach x Sekunden (siehe F10)"
    Label7.Caption = LoadResString(1210 + Sprache)  '"F7 - Stoppe Automatik"
    Label8.Caption = LoadResString(1211 + Sprache)  '"F8 - Neue Diashow zusammenstellen"
    Label9.Caption = LoadResString(1212 + Sprache)  '"F9 - Mauszeiger sichtbar/unsichtbar"
    Label10.Caption = LoadResString(1213 + Sprache)  '"F10 - Dialog-Fenster zum Einstellen von Werten
    Label11.Caption = LoadResString(1214 + Sprache)  '"F11 - zeigt Bildgröße und Position des Mauszeigers"
    LabelStrgC.Caption = LoadResString(2349 + Sprache)  '"Strg+C - kopiert das Bild"                        'Gerbing 12.08.2017
    lblStrgG.Caption = LoadResString(1129 + Sprache)    '"Strg+G - zeige die GEO-Position
    LabelStrgZ.Caption = LoadResString(1215 + Sprache)  '"Strg+Z - macht die Rechteck-Lupe scharf, vorher ist sie nicht benutzbar"
    LabelStrgO.Caption = LoadResString(1216 + Sprache)  '"Strg+O - schaltet die Rechteck-Lupe wieder aus"
    lblAltF4.Caption = LoadResString(1217 + Sprache)  '"Alt+F4 - Beenden"
    
    'SetWindowPos Me.hWnd, HWND_TOPMOST, 200, 200, 450, 400, SWP_NOACTIVATE Or SWP_SHOWWINDOW
    'SetWindowPos Me.hWnd, HWND_TOPMOST, 200, 200, 650, 400, SWP_NOACTIVATE Or SWP_SHOWWINDOW                'Gerbing 16.03.2009
    SetWindowPos Me.hwnd, HWND_TOPMOST, 200, 200, 850, 600, SWP_NOACTIVATE Or SWP_SHOWWINDOW                'Gerbing 03.03.2012
End Sub

Private Sub Label1_Click()
    KeyCode = vbKeyF1
    Shift = 0
    UnloadMe
    Call DiashowForm.Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub Label11_Click()
    XYPos.Show
End Sub

Private Sub Label2_Click()
    KeyCode = vbKeyF2
    Shift = 0
    UnloadMe
    Call DiashowForm.Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub Label3_Click()
    KeyCode = vbKeyF3
    Shift = 0
    UnloadMe
    Call DiashowForm.Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub Label4_Click()
    KeyCode = vbKeyF4
    Shift = 0
    UnloadMe
    Call DiashowForm.Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub Label5_Click()
    ListBoxForm.Show
    ListBoxForm.ZOrder 0
    'ListBoxForm.chkIptcAnzeigen.Value = 1                                      'Gerbing 09.06.2014
End Sub

Private Sub Label6_Click()
    DiashowForm.F6TimerGestartet = True
    DiashowForm.Timer1.Enabled = True
    frmBildMitGDIPlus.Show
End Sub

Private Sub Label7_Click()
    DiashowForm.F6TimerGestartet = False
    DiashowForm.Timer1.Enabled = False
    frmBildMitGDIPlus.Show
End Sub

Private Sub Label8_Click()
    KeyCode = vbKeyF8
    Shift = 0
    UnloadMe
    Call DiashowForm.Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub Label9_Click()
    KeyCode = vbKeyF9
    Shift = 0
    UnloadMe
    Call frmBildMitGDIPlus.Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub Label10_Click()
    WertxForm.Show
    WertxForm.ZOrder 0
End Sub

Private Sub LabelStrgC_Click()                                          'Gerbing 11.08.2017
    KeyCode = vbKeyC
    Shift = vbCtrlMask
    UnloadMe
    Call frmBildMitGDIPlus.Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub LabelStrgO_Click()
    KeyCode = vbKeyO
    Shift = vbCtrlMask
    UnloadMe
    Call frmBildMitGDIPlus.Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub LabelStrgZ_Click()
    KeyCode = vbKeyZ
    Shift = vbCtrlMask
    UnloadMe
    Call frmBildMitGDIPlus.Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub lblAlt_Click()                                              'Gerbing 29.01.2018
    KeyCode = 18                                                        '18 = Menu key
    Shift = vbAltMask                                                   'Alt
    UnloadMe
    Call frmBildMitGDIPlus.Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub lblAltF4_Click()
    End
End Sub

Private Sub UnloadMe()
    'Unload Me
    Me.Hide
    Unload frmBildMitGDIPlus
End Sub

Private Sub lblStrgG_Click()                                            'Gerbing 29.01.2018
'    KeyCode = vbKeyG
'    Shift = vbCtrlMask
'    UnloadMe
'    Call frmBildMitGDIPlus.Form_KeyDown(KeyCode, Shift)
    Call ListBoxForm.ZeigeGeoPosition
End Sub
