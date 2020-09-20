VERSION 5.00
Begin VB.Form frmFotoAlbumWirdGeladen 
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   1320
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   6864
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   110
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   572
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   432
      Left            =   120
      Picture         =   "frmFotoAlbumWirdGeladen.frx":0000
      ScaleHeight     =   263.118
      ScaleMode       =   0  'Benutzerdefiniert
      ScaleWidth      =   263.118
      TabIndex        =   1
      Top             =   120
      Width           =   432
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fest Einfach
      Height          =   432
      Left            =   600
      Picture         =   "frmFotoAlbumWirdGeladen.frx":0C42
      Top             =   120
      Width           =   6156
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      Caption         =   "GERBING Fotoalbum wird geladen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6612
   End
End
Attribute VB_Name = "frmFotoAlbumWirdGeladen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    
    ' SetWindowPos Flags
    Const SWP_NOSIZE = &H1
    Const SWP_NOMOVE = &H2
    Const SWP_NOZORDER = &H4
    Const SWP_NOREDRAW = &H8
    Const SWP_NOACTIVATE = &H10
    Const SWP_DRAWFRAME = &H20
    Const SWP_SHOWWINDOW = &H40
    Const SWP_HIDEWINDOW = &H80
    Const SWP_NOCOPYBITS = &H100
    Const SWP_NOREPOSITION = &H200

    ' SetWindowPos() hwndInsertAfter values
    Const HWND_TOP = 0
    Const HWND_BOTTOM = 1
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2

Private Sub Form_Load()
    Dim HWNDInteger As Long
    Dim TLeft As Long
    Dim TTop As Long
    Dim TWidth As Long
    Dim THeight As Long

    Label1.Caption = LoadResString(1848 + Sprache)                      'GERBING Fotoalbum wird geladen
    screenWidth = GetDeviceCaps(Me.hDC, HORZRES)
    screenHeight = GetDeviceCaps(Me.hDC, VERTRES)

    HWNDInteger = HWND_TOPMOST
'    TLeft = 100
'    TTop = 100
'    TWidth = 900
'    THeight = 200

    TLeft = screenWidth \ 2 - (Me.Width \ Screen.TwipsPerPixelX) \ 2
    TTop = screenHeight \ 2 - (Me.Height \ Screen.TwipsPerPixelY) \ 2
    TWidth = Me.Width \ Screen.TwipsPerPixelX
    THeight = Me.Height \ Screen.TwipsPerPixelY

    SetWindowPos Me.hWnd, HWNDInteger, TLeft, TTop, TWidth, THeight, SWP_SHOWWINDOW Or SWP_NOACTIVATE
    Screen.MousePointer = vbHourglass
End Sub
