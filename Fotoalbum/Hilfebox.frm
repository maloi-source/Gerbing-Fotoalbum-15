VERSION 5.00
Begin VB.Form Hilfebx 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Hilfe zu den Bedien-Tasten für Fotos"
   ClientHeight    =   9612
   ClientLeft      =   2832
   ClientTop       =   2856
   ClientWidth     =   9888
   Icon            =   "Hilfebox.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   9612
   ScaleWidth      =   9888
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Label LabelAlt 
      Caption         =   "Alt = Menü ein/aus"
      Height          =   372
      Left            =   120
      TabIndex        =   24
      Top             =   840
      Width           =   8532
   End
   Begin VB.Label LabelStrg_C 
      Caption         =   "Strg + C =  kopiert das Bild"
      Height          =   252
      Left            =   120
      TabIndex        =   23
      Top             =   6240
      Width           =   8412
   End
   Begin VB.Label lblStrg_G 
      Caption         =   "Strg + G = Zeige die GEO-Position"
      Height          =   252
      Left            =   120
      TabIndex        =   22
      Top             =   6600
      Width           =   9492
   End
   Begin VB.Label LabelNumStrgM 
      Caption         =   "LabelNumStrgM"
      Height          =   372
      Left            =   120
      TabIndex        =   21
      Top             =   8760
      Width           =   8412
   End
   Begin VB.Label LabelNumStrgN 
      Caption         =   "LabelNumStrgN"
      Height          =   372
      Left            =   120
      TabIndex        =   20
      Top             =   8400
      Width           =   8412
   End
   Begin VB.Label Label24 
      Caption         =   "Umsch + F5 = Zeilenfenster der Stichworte"
      Height          =   372
      Left            =   120
      TabIndex        =   19
      Top             =   3000
      Width           =   8292
   End
   Begin VB.Label Label20 
      Caption         =   "Strg + B   = Bildgröße und Position des Mauszeigers liefern"
      Height          =   372
      Left            =   120
      TabIndex        =   18
      Top             =   5880
      Width           =   8292
   End
   Begin VB.Label Label19 
      Caption         =   "Strg + O  = schaltet die Rechteck-Lupe wieder aus"
      Height          =   372
      Left            =   120
      TabIndex        =   17
      Top             =   8040
      Width           =   8292
   End
   Begin VB.Label Label18 
      Caption         =   "Strg + Z   = macht die Rechteck-Lupe scharf, vorher ist sie nicht benutzbar"
      Height          =   372
      Left            =   120
      TabIndex        =   16
      Top             =   7680
      Width           =   9612
   End
   Begin VB.Label Label14 
      Caption         =   "Bei Video im Vollbild-Modus wirken keine Funktionstasten. Sie müssen erst den Vollbild-Modus wieder ausschalten."
      Height          =   852
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   8292
   End
   Begin VB.Label Label17 
      Caption         =   "Strg + I    = Import mit Berücksichtigung der Jahres-Unterverzeichnisse "
      Height          =   372
      Left            =   120
      TabIndex        =   14
      Top             =   7320
      Width           =   9612
   End
   Begin VB.Label Label15 
      Caption         =   "Strg + K   = Öffnet das Fenster zum Export der aktuell selektierten Fotos "
      Height          =   372
      Left            =   120
      TabIndex        =   13
      Top             =   6960
      Width           =   9492
   End
   Begin VB.Label Label9 
      Caption         =   "F9 = Mauszeiger sichtbar/unsichtbar"
      Height          =   372
      Left            =   120
      TabIndex        =   12
      Top             =   4440
      Width           =   8292
   End
   Begin VB.Label Label13 
      Caption         =   "F12 = Dialog-Fenster zum Einstellen von Werten"
      Height          =   372
      Left            =   120
      TabIndex        =   11
      Top             =   5520
      Width           =   8292
   End
   Begin VB.Label Label12 
      Caption         =   "F11 = Kommentar-Fenster ausblenden"
      Height          =   372
      Left            =   120
      TabIndex        =   10
      Top             =   5160
      Width           =   8292
   End
   Begin VB.Label Label11 
      Caption         =   "F10 = Kommentar-Fenster einblenden, falls das Kommentar-Feld nicht leer ist"
      Height          =   372
      Left            =   120
      TabIndex        =   9
      Top             =   4800
      Width           =   9492
   End
   Begin VB.Label Label10 
      Caption         =   "Alt + F4 = Programm beenden"
      Height          =   372
      Left            =   120
      TabIndex        =   8
      Top             =   9120
      Width           =   8292
   End
   Begin VB.Label Label8 
      Caption         =   "F7 = Stoppe Automatik"
      Height          =   372
      Left            =   120
      TabIndex        =   7
      Top             =   3720
      Width           =   8292
   End
   Begin VB.Label Label7 
      Caption         =   "Strg + F6 = Starte Automatik alle x Sekunden (siehe F12) bzw Videos kontinuierlich"
      Height          =   372
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   9492
   End
   Begin VB.Label Label6 
      Caption         =   "F8 = Startet neue Suche in der Datenbank / Programm beenden"
      Height          =   372
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   8292
   End
   Begin VB.Label Label5 
      Caption         =   "F5 = Listenfenster der Stichworte"
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   8292
   End
   Begin VB.Label Label4 
      Caption         =   "F4 = Zoom In (Bild vergrößern)"
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   8292
   End
   Begin VB.Label Label3 
      Caption         =   "F3 = Gehe ein Bild vorwärts (oder Taste Alt und ->)"
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   8412
   End
   Begin VB.Label Label2 
      Caption         =   "F2 = Gehe ein Bild rückwärts (oder Taste Alt und <-)"
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   8412
   End
   Begin VB.Label Label1 
      Caption         =   "F1 = Zoom Out (Bild verkleinern)"
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   8412
   End
End
Attribute VB_Name = "Hilfebx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim KeyCode As Integer
    Dim Shift As Integer
    Dim blnHilfebxRightButton As Boolean                            'Gerbing 15.02.2013

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RetteKeyCode As Integer
    
    If gblnComefromVideo = True Then
        If KeyCode = vbKeyF1 Or KeyCode = vbKeyF4 Then Exit Sub     'Gerbing 26.11.2012
    End If
    '----------------------------------------------------------------------------------
    RetteKeyCode = KeyCode                                          'Gerbing 11.08.2012
    Call Form1.Form_KeyDown(KeyCode, Shift)
    MeHide
End Sub

Private Sub Form_Load()
'BorderStyle = 4 Festes Werkzeugfenster einstellen
    
    Call AnpassenNutzerWunsch(Me)                               'Gerbing 11.03.2017
'    If Query.chkFensterGrößeÄnderbar.Value = 1 Then             'Gerbing 06.12.2005 01.09.2017
'        Me.Top = Form1.Top                                      'Gerbing 06.12.2006
'        Me.Left = Form1.Left
'    End If
    
    'Me.Caption = "Hilfe zu den Bedien-Tasten für Fotos"        'Gerbing 08.11.2005
    Me.Caption = LoadResString(1051 + Sprache)
    Label1.Caption = LoadResString(1052 + Sprache)            'F1
    Label2.Caption = LoadResString(1053 + Sprache)            'F2
    Label3.Caption = LoadResString(1054 + Sprache)            'F3
    Label4.Caption = LoadResString(1055 + Sprache)            'F4
    Label5.Caption = LoadResString(1056 + Sprache)            'F5
    Label6.Caption = LoadResString(1057 + Sprache)            'F8
    Label7.Caption = LoadResString(1058 + Sprache)            'Strg+F6
    Label8.Caption = LoadResString(1059 + Sprache)            'F7
    Label9.Caption = LoadResString(1060 + Sprache)            'F9
    Label10.Caption = LoadResString(1061 + Sprache)           'Alt+F4
    Label11.Caption = LoadResString(1062 + Sprache)           'F10
    Label12.Caption = LoadResString(1063 + Sprache)           'F11
    Label13.Caption = LoadResString(1064 + Sprache)           'F12
    Label14.Caption = LoadResString(1065 + Sprache)           'Bei Video im Vollbildmodus wirken keine Funktionstasten
    LabelStrg_C.Caption = LoadResString(2349 + Sprache)       '"Strg+C - kopiert das Bild"                      'Gerbing 11.08.2017
    lblStrg_G.Caption = LoadResString(1129 + Sprache)         'Strg+G                                           'Gerbing 03.10.2016
    Label15.Caption = LoadResString(1066 + Sprache)           'Strg+K
    Label17.Caption = LoadResString(1068 + Sprache)           'Strg+I
    Label18.Caption = LoadResString(1069 + Sprache)           'Strg+Z
    Label19.Caption = LoadResString(1070 + Sprache)           'Strg+O
    Label20.Caption = LoadResString(1071 + Sprache)           'Strg+B
    Label24.Caption = LoadResString(1075 + Sprache)           'Umsch+F5
    LabelAlt.Caption = LoadResString(3111 + Sprache)          'Alt
    LabelNumStrgN.Caption = LoadResString(1236 + Sprache) 'Num+Strg+N - Zeile Bildbeschreibung einschalten      'Gerbing 03.03.2012
    LabelNumStrgM.Caption = LoadResString(1237 + Sprache) 'Num+Strg+M - Zeile Bildbeschreibung ausschalten      'Gerbing 03.03.2012
    
    'SetWindowPos hwnd, conHwndTopmost, 200, 150, 450, 450, conSwpNoActivate Or conSwpShowWindow
    Hilfebx.KeyPreview = True
End Sub

'
'Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'Gerbing 05.02.2013 auskommentiert
'    MeHide
'End Sub


Private Sub Form_Unload(Cancel As Integer)
    If gblnComefromVideo = True Then                                                    'Gerbing 16.06.2012
        frmVideo.Show
    End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblnComefromVideo = True Or Button = vbRightButton Then Exit Sub                 'Gerbing 26.11.2012 Gerbing 15.02.2013
    
    KeyCode = vbKeyF1                                               'F1
    Shift = 0
    Call Form1.Form_KeyDown(KeyCode, Shift)                         'F1
    MeHide                                                                              'Gerbing 04.12.2012
End Sub

Private Sub Label10_Click()                                         'Alt+F4
    If blnHilfebxRightButton = True Then Exit Sub                                       'Gerbing 15.02.2013
    
    Unload frmGridAndThumb
    Unload Hilfebx
    Unload KommentarForm
    Unload Query
    'Unload QueryJedesFeld
    Unload MP
    End
End Sub

Private Sub Label10_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)  ' Alt+F4 'Gerbing 15.02.2013
    If Button = vbRightButton Then
        blnHilfebxRightButton = True
    Else
        blnHilfebxRightButton = False
    End If
End Sub

Private Sub Label11_Click()
    If blnHilfebxRightButton = True Then Exit Sub                                       'Gerbing 15.02.2013
    
    KeyCode = vbKeyF10                                              'F10
    Shift = 0
    Call Form1.Form_KeyDown(KeyCode, Shift)
    If gblnComefromVideo = True Then                                                    'Gerbing 16.06.2012
        frmVideo.Show
    End If
    MeHide                                                                              'Gerbing 04.12.2012
End Sub

Private Sub Label11_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'F10 'Gerbing 15.02.2013
    If Button = vbRightButton Then
        blnHilfebxRightButton = True
    Else
        blnHilfebxRightButton = False
    End If
End Sub

Private Sub Label12_Click()                                                             'F11
    If blnHilfebxRightButton = True Then Exit Sub                                       'Gerbing 15.02.2013
    
    KeyCode = vbKeyF11                                              'F11
    Shift = 0
    Call Form1.Form_KeyDown(KeyCode, Shift)
    If gblnComefromVideo = True Then                                                    'Gerbing 16.06.2012
        frmVideo.Show
    End If
    MeHide                                                                              'Gerbing 04.12.2012
End Sub

Private Sub Label12_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'F11 'Gerbing 15.02.2013
    If Button = vbRightButton Then
        blnHilfebxRightButton = True
    Else
        blnHilfebxRightButton = False
    End If
End Sub

Private Sub Label13_Click()
    If blnHilfebxRightButton = True Then Exit Sub                                       'Gerbing 15.02.2013
    
    KeyCode = vbKeyF12                                              'F12
    Shift = 0
    Call Form1.Form_KeyDown(KeyCode, Shift)
    If gblnComefromVideo = True Then                                                    'Gerbing 16.06.2012
        frmVideo.Show
    End If
    MeHide                                                                              'Gerbing 04.12.2012
End Sub

Private Sub Label13_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'Gerbing 15.02.2013
    If Button = vbRightButton Then
        blnHilfebxRightButton = True
    Else
        blnHilfebxRightButton = False
    End If
End Sub

Private Sub Label15_Click()
    If blnHilfebxRightButton = True Then Exit Sub                                       'Gerbing 15.02.2013
    
    KeyCode = vbKeyK
    Shift = vbCtrlMask                                              'Strg+K gleichzeitig
    Call Form1.Form_KeyDown(KeyCode, Shift)
    MeHide                                                                              'Gerbing 04.12.2012
End Sub

Private Sub Label15_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'Gerbing 15.02.2013
    If Button = vbRightButton Then
        blnHilfebxRightButton = True
    Else
        blnHilfebxRightButton = False
    End If
End Sub

Private Sub Label17_Click()
    If blnHilfebxRightButton = True Then Exit Sub                                       'Gerbing 15.02.2013
    
    KeyCode = vbKeyI
    Shift = vbCtrlMask                                              'Strg+I gleichzeitig
    Call Form1.Form_KeyDown(KeyCode, Shift)
    MeHide                                                                              'Gerbing 04.12.2012
End Sub

Private Sub Label17_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'Gerbing 15.02.2013
    If Button = vbRightButton Then
        blnHilfebxRightButton = True
    Else
        blnHilfebxRightButton = False
    End If
End Sub

Private Sub Label18_Click()
    If blnHilfebxRightButton = True Then Exit Sub                                       'Gerbing 15.02.2013
    If gblnComefromVideo = True Then Exit Sub                                           'Gerbing 07.05.2013

    If gblnComefromVideo = True Then                                                    'Gerbing 16.06.2012
        Me.Hide
        Exit Sub
    End If
    KeyCode = vbKeyZ
    Shift = vbCtrlMask                                              'Strg+Z gleichzeitig
    Call Form1.Form_KeyDown(KeyCode, Shift)
    MeHide                                                                              'Gerbing 04.12.2012
End Sub

Private Sub Label18_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'Gerbing 15.02.2013
    If Button = vbRightButton Then
        blnHilfebxRightButton = True
    Else
        blnHilfebxRightButton = False
    End If
End Sub

Private Sub Label19_Click()
    If blnHilfebxRightButton = True Then Exit Sub                                       'Gerbing 15.02.2013
    If gblnComefromVideo = True Then Exit Sub                                           'Gerbing 07.05.2013

    If gblnComefromVideo = True Then                                                    'Gerbing 16.06.2012
        Me.Hide
        Exit Sub
    End If
    KeyCode = vbKeyO
    Shift = vbCtrlMask                                              'Strg+O gleichzeitig
    Call Form1.Form_KeyDown(KeyCode, Shift)
    MeHide                                                                              'Gerbing 04.12.2012
End Sub

Private Sub Label19_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'Gerbing 15.02.2013
    If Button = vbRightButton Then
        blnHilfebxRightButton = True
    Else
        blnHilfebxRightButton = False
    End If
End Sub

Private Sub Label2_Click()
    If blnHilfebxRightButton = True Then Exit Sub                                       'Gerbing 15.02.2013
    
    KeyCode = vbKeyF2                                               'F2
    Shift = 0
    Call Form1.Form_KeyDown(KeyCode, Shift)
    MeHide                                                                              'Gerbing 04.12.2012
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'Gerbing 15.02.2013
    If Button = vbRightButton Then
        blnHilfebxRightButton = True
    Else
        blnHilfebxRightButton = False
    End If
End Sub

Private Sub Label20_Click()
    If blnHilfebxRightButton = True Then Exit Sub                                       'Gerbing 15.02.2013
    If gblnComefromVideo = True Then Exit Sub                                           'Gerbing 07.05.2013

    If gblnComefromVideo = True Then                                                    'Gerbing 16.06.2012
        Me.Hide
        Exit Sub
    End If
    KeyCode = vbKeyB
    Shift = vbCtrlMask                                              'Strg+B gleichzeitig
    Call Form1.Form_KeyDown(KeyCode, Shift)
    MeHide                                                                              'Gerbing 04.12.2012
    gblnShowXYPos = True                                                                'Gerbing 04.12.2012
End Sub

Private Sub Label20_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'Gerbing 15.02.2013
    If Button = vbRightButton Then
        blnHilfebxRightButton = True
    Else
        blnHilfebxRightButton = False
    End If
End Sub

Private Sub Label23_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'Gerbing 15.02.2013
    If Button = vbRightButton Then
        blnHilfebxRightButton = True
    Else
        blnHilfebxRightButton = False
    End If
End Sub

Private Sub Label24_Click() 'Gerbing 30.06.2005
    If blnHilfebxRightButton = True Then Exit Sub                                       'Gerbing 15.02.2013

    KeyCode = vbKeyF5
    Shift = vbShiftMask                                             'Umsch + F5 gleichzeitig
    Call Form1.Form_KeyDown(KeyCode, Shift)
    MeHide                                                                              'Gerbing 04.12.2012
End Sub

Private Sub Label24_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'Gerbing 15.02.2013
    If Button = vbRightButton Then
        blnHilfebxRightButton = True
    Else
        blnHilfebxRightButton = False
    End If
End Sub

Private Sub Label3_Click()
    If blnHilfebxRightButton = True Then Exit Sub                                       'Gerbing 15.02.2013

    KeyCode = vbKeyF3                                               'F3
    Shift = 0
    Call Form1.Form_KeyDown(KeyCode, Shift)
    MeHide                                                                              'Gerbing 04.12.2012
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'Gerbing 15.02.2013
    If Button = vbRightButton Then
        blnHilfebxRightButton = True
    Else
        blnHilfebxRightButton = False
    End If
End Sub

Private Sub Label4_Click()
    If gblnComefromVideo = True Or blnHilfebxRightButton = True Then Exit Sub                  'Gerbing 26.11.2012 Gerbing 15.02.2013
    
    KeyCode = vbKeyF4                                               'F4
    Shift = 0
    Call Form1.Form_KeyDown(KeyCode, Shift)
    MeHide                                                                              'Gerbing 04.12.2012
End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'Gerbing 15.02.2013
    If Button = vbRightButton Then
        blnHilfebxRightButton = True
    Else
        blnHilfebxRightButton = False
    End If
End Sub

Private Sub Label5_Click()                                          'F5
    If blnHilfebxRightButton = True Then Exit Sub                                       'Gerbing 15.02.2013
    
    frmGridAndThumb.Show
End Sub

Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'Gerbing 15.02.2013
    If Button = vbRightButton Then
        blnHilfebxRightButton = True
    Else
        blnHilfebxRightButton = False
    End If
End Sub

Private Sub Label6_Click()
    If blnHilfebxRightButton = True Then Exit Sub                                       'Gerbing 15.02.2013
    
    KeyCode = vbKeyF8                                               'F8
    Shift = 0
    Call Form1.Form_KeyDown(KeyCode, Shift)
    Me.Hide                                                                             'Gerbing 04.12.2012
End Sub

Private Sub Label6_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        blnHilfebxRightButton = True
    Else
        blnHilfebxRightButton = False
    End If
End Sub

Private Sub Label7_Click()
    If blnHilfebxRightButton = True Then Exit Sub                                       'Gerbing 15.02.2013
    
    KeyCode = vbKeyF6
    Shift = vbCtrlMask      'Gerbing 03.11.2004                     Strg+F6 gleichzeitig
    Call Form1.Form_KeyDown(KeyCode, Shift)
    If gblnComefromVideo = True Then                                                    'Gerbing 16.06.2012
        frmVideo.Show
    End If
    MeHide                                                                              'Gerbing 04.12.2012
End Sub

Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'Gerbing 15.02.2013
    If Button = vbRightButton Then
        blnHilfebxRightButton = True
    Else
        blnHilfebxRightButton = False
    End If
End Sub

Private Sub Label8_Click()
    If blnHilfebxRightButton = True Then Exit Sub                                       'Gerbing 15.02.2013
    
    KeyCode = vbKeyF7                                               'F7
    Shift = 0
    Call Form1.Form_KeyDown(KeyCode, Shift)
    If gblnComefromVideo = True Then                                                    'Gerbing 16.06.2012
        frmVideo.Show
    End If
    MeHide                                                                              'Gerbing 04.12.2012
End Sub

Private Sub Label8_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'Gerbing 15.02.2013
    If Button = vbRightButton Then
        blnHilfebxRightButton = True
    Else
        blnHilfebxRightButton = False
    End If
End Sub

Private Sub Label9_Click()
    If blnHilfebxRightButton = True Then Exit Sub                                       'Gerbing 15.02.2013
    If gblnComefromVideo = True Then Exit Sub                                           'Gerbing 07.05.2013
    
    KeyCode = vbKeyF9                                               'F9
    Shift = 0
    Call Form1.Form_KeyDown(KeyCode, Shift)
    If gblnComefromVideo = True Then                                                    'Gerbing 16.06.2012
        frmVideo.Show
    End If
    MeHide                                                                              'Gerbing 04.12.2012
End Sub

Private Sub MeHide()
    Me.Hide
End Sub

Private Sub Label9_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'Gerbing 15.02.2013
    If Button = vbRightButton Then
        blnHilfebxRightButton = True
    Else
        blnHilfebxRightButton = False
    End If
End Sub

Private Sub LabelAlt_Click()                                                            'Gerbing 01.09.2017
    KeyCode = 18                                                    '18 = Menu key
    Shift = vbAltMask                                               'Alt
    If gblnComefromVideo = True Then
        Call frmVideo.Form_KeyDown(KeyCode, Shift)                                      'Gerbing 01.09.2017
    Else
        Call Form1.Form_KeyDown(KeyCode, Shift)
    End If
    MeHide                                                                              'Gerbing 04.12.2012
End Sub

Private Sub LabelNumStrgM_Click()
    If blnHilfebxRightButton = True Then Exit Sub                                       'Gerbing 15.02.2013
    
    KeyCode = vbKeyM
    Shift = vbCtrlMask + vbShiftMask                                'Strg+Num+M gleichzeitig
    Call Form1.Form_KeyDown(KeyCode, Shift)
    If gblnComefromVideo = True Then                                                    'Gerbing 16.06.2012
        frmVideo.Show
    End If
    MeHide                                                                              'Gerbing 04.12.2012
End Sub

Private Sub LabelNumStrgM_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'Gerbing 15.02.2013
    If Button = vbRightButton Then
        blnHilfebxRightButton = True
    Else
        blnHilfebxRightButton = False
    End If
End Sub

Private Sub LabelNumStrgN_Click()
    If blnHilfebxRightButton = True Then Exit Sub                                       'Gerbing 15.02.2013
    
    KeyCode = vbKeyN
    Shift = vbCtrlMask + vbShiftMask                                'Strg+Num+N gleichzeitig
    Call Form1.Form_KeyDown(KeyCode, Shift)
    If gblnComefromVideo = True Then                                                    'Gerbing 16.06.2012
        frmVideo.Show
    End If
    MeHide                                                                              'Gerbing 04.12.2012
End Sub

Private Sub LabelNumStrgN_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'Gerbing 15.02.2013
    If Button = vbRightButton Then
        blnHilfebxRightButton = True
    Else
        blnHilfebxRightButton = False
    End If
End Sub

Private Sub LabelStrg_C_Click()                                                         'Gerbing 11.08.2017
    'If blnHilfebxRightButton = True Then Exit Sub
    
    KeyCode = vbKeyC                                                'Strg+C gleichzeitig
    Shift = vbCtrlMask
'    If gblnComefromVideo = True Then
'        Call frmVideo.Form_KeyDown(KeyCode, Shift)
'    Else
        Call Form1.Form_KeyDown(KeyCode, Shift)
'    End If
    MeHide
End Sub

Private Sub lblStrg_G_Click()                                                           'Gerbing 03.10.2016
    If blnHilfebxRightButton = True Then Exit Sub
    
    KeyCode = vbKeyG
    Shift = vbCtrlMask                                              'Strg+G gleichzeitig
    Call Form1.Form_KeyDown(KeyCode, Shift)
    MeHide
End Sub

Private Sub lblStrg_G_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'Gerbing 03.10.2016
    If Button = vbRightButton Then
        blnHilfebxRightButton = True
    Else
        blnHilfebxRightButton = False
    End If
End Sub
