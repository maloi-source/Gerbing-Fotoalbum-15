VERSION 5.00
Begin VB.Form frmGEOPosition 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "GEO position"
   ClientHeight    =   7752
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   8088
   Icon            =   "frmGEOPosition.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7752
   ScaleWidth      =   8088
   StartUpPosition =   1  'Fenstermitte
   Begin VB.VScrollBar VScroll1 
      Height          =   1500
      Left            =   840
      Max             =   20
      TabIndex        =   3
      Top             =   2280
      Value           =   16
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.ComboBox cmbZoom 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   120
      Style           =   2  'Dropdown-Liste
      TabIndex        =   2
      Top             =   4080
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.ComboBox cmbMapType 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   600
      Style           =   2  'Dropdown-Liste
      TabIndex        =   1
      Top             =   550
      Width           =   2300
   End
   Begin VB.TextBox txtCenter 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2772
   End
   Begin FotoAlbum.ucGMap ucGMap1 
      Height          =   7680
      Left            =   480
      TabIndex        =   4
      Top             =   0
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   13547
   End
   Begin VB.Image imgMinus 
      Height          =   432
      Left            =   0
      Picture         =   "frmGEOPosition.frx":038A
      Stretch         =   -1  'True
      Top             =   480
      Width           =   432
   End
   Begin VB.Image imgPlus 
      Height          =   432
      Left            =   0
      Picture         =   "frmGEOPosition.frx":0714
      Stretch         =   -1  'True
      Top             =   840
      Width           =   432
   End
End
Attribute VB_Name = "frmGEOPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
 
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1, vbKeyF2, vbKeyF3, vbKeyF4                                 'Gerbing 03.10.2016
            'Debug.Print "Form_KeyDown-Case vbKeyF1, vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF8, vbKeyF10"
            Unload Me
            'Tastatur-Eingabe weiterreichen
            '-> Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
                Call Form1.Form_KeyDown(KeyCode, Shift)
    End Select
End Sub

Private Sub Form_Load()
    Dim i&
    Dim GMouseCoordLatLng As String
    Dim MarkerChar As String
        
    Call AnpassenNutzerWunsch(Me)                                               'Gerbing 11.03.2017
    'Me.Caption = "GEO-Position anzeigen"                                       'Gerbing 05.09.2016
    Me.Caption = LoadResString(3162 + Sprache)
    For i = 0 To 20: cmbZoom.AddItem 2 ^ i: Next
    cmbZoom.ListIndex = 16
    
    For i = 0 To 3: cmbMapType.AddItem ucGMap1.GetMapType(i): Next
    cmbMapType.ListIndex = 0
    
    'gstrGEOPosition = "50.83266,12.81863"  'Chemnitz Georg-Weerth-Straße
    'GMouseCoordLatLng = "50.83266,12.81863"
    GMouseCoordLatLng = gstrGEOPosition
    MarkerChar = Chr$(65)
    ucGMap1.AddMarker GMouseCoordLatLng, vbGreen, MarkerChar
    
    ucGMap1.GPoint = Mid$(gstrGEOPosition, 1)
    txtCenter.Text = "GEO position: " & ucGMap1.GPoint
    'txtSexagesimal.Text = frmGridAndThumb.GPSLatitudeRef & " " & frmGridAndThumb.GPSLatitude & " , " & frmGridAndThumb.GPSLongitudeRef & " " & frmGridAndThumb.GPSLongitude
End Sub
 
Private Sub cmbZoom_Click()
    ucGMap1.GZoom = cmbZoom.ListIndex
End Sub

Private Sub cmbMapType_Click()
    ucGMap1.MapType = cmbMapType.ListIndex
End Sub

Private Sub ucGMap1_DblClick(ByVal GMouseCoordLatLng As String)
    ucGMap1.GPoint = Mid$(GMouseCoordLatLng, 1)                         'Gerbing 02.09.2016 auf Doppelklick wird zentriert
End Sub

'Private Sub ucGMap1_MouseMove(ByVal GMouseCoordLatLng As String)
'    txtMouseLatLng.Text = "MousePoint: " & GMouseCoordLatLng
'End Sub
Private Sub ucGMap1_MouseUp(ByVal GMouseCoordLatLng As String, Button As Integer, Shift As Integer, x As Single, y As Single)
    'cmbZoom.SetFocus
End Sub

Private Sub imgMinus_Click()
    If VScroll1.Value > 0 Then
        VScroll1.Value = VScroll1.Value - 1
    End If
End Sub

Private Sub imgPlus_Click()
    If VScroll1.Value < 20 Then
        VScroll1.Value = VScroll1.Value + 1
    End If
End Sub

Private Sub VScroll1_Change()
    cmbZoom.ListIndex = VScroll1.Value
End Sub

