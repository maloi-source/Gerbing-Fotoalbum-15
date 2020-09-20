VERSION 5.00
Object = "{B6CC61F6-3F1A-4B00-9918-13F66F185263}#1.0#0"; "lblctlsu.ocx"
Object = "{A10D6B26-9A8F-4A87-A2D1-1D8C9EED0967}#1.3#0"; "statbaru.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGridAndThumb 
   BackColor       =   &H00800000&
   ClientHeight    =   11400
   ClientLeft      =   192
   ClientTop       =   516
   ClientWidth     =   18228
   Icon            =   "frmGridAndThumb.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MousePointer    =   7  'Größenänderung N S
   ScaleHeight     =   950
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1519
   Begin VB.PictureBox pbBottomLeer 
      BackColor       =   &H00C0C0C0&
      Height          =   2532
      Left            =   7920
      ScaleHeight     =   207
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   687
      TabIndex        =   11
      Top             =   4920
      Width           =   8292
      Begin VB.Label lblKeineThumbnailsAusgewählt 
         BackColor       =   &H00C0C0C0&
         Caption         =   "es sind keine Thumbnails ausgewählt oder es sind keine vorhanden"
         Height          =   732
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   8052
      End
   End
   Begin VB.PictureBox pbTop 
      BackColor       =   &H00C0C0C0&
      Height          =   6132
      Left            =   0
      MousePointer    =   1  'Pfeil
      ScaleHeight     =   507
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1447
      TabIndex        =   3
      Top             =   0
      Width           =   17412
      Begin VB.CommandButton btnThumbnailsAbbrechen 
         Caption         =   "Thumbnails a&bbrechen"
         Height          =   372
         Left            =   8640
         TabIndex        =   10
         Top             =   120
         Width           =   2772
      End
      Begin VB.CommandButton btnMitThumbnails 
         Caption         =   "mit &Thumbnails"
         Height          =   372
         Left            =   5640
         TabIndex        =   9
         Top             =   120
         Width           =   2772
      End
      Begin VB.CommandButton btnShowUsers 
         Caption         =   "Show users"
         Height          =   492
         Left            =   11520
         TabIndex        =   8
         Top             =   3720
         Visible         =   0   'False
         Width           =   3252
      End
      Begin VB.CommandButton btnRefresh 
         Height          =   495
         Left            =   11640
         Picture         =   "frmGridAndThumb.frx":038A
         Style           =   1  'Grafisch
         TabIndex        =   7
         ToolTipText     =   "Aktualisieren - nur sinnvoll in Multiuser-Umgebung"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton btnSpaltenbreitenSpeichern 
         Caption         =   "&Spaltenbreiten speichern"
         Height          =   372
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Sie können die Spaltenbreite mit der Maus durch Ziehen verändern. Diese Einstellung wird hiermit gespeichert."
         Top             =   120
         Width           =   5292
      End
      Begin VB.CommandButton btnMerkerspalteEinschalten 
         Caption         =   "&Merkerspalte in jedem Datensatz ein/ausschalten"
         Height          =   372
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Mit der Merkerspalte können Sie Fotos vormerken, die Sie exportieren oder löschen oder weiterselektieren wollen"
         Top             =   600
         Width           =   5292
      End
      Begin VB.CommandButton btnOhneThumbnails 
         Caption         =   "&ohne Thumbnails"
         Height          =   372
         Left            =   5640
         TabIndex        =   4
         Top             =   600
         Width           =   2772
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   372
         Left            =   8040
         Top             =   1320
         Visible         =   0   'False
         Width           =   2052
         _ExtentX        =   3620
         _ExtentY        =   656
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   """Fotos"""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DBGridNeu 
         Height          =   3612
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   6972
         _ExtentX        =   12298
         _ExtentY        =   6371
         _Version        =   393216
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   19
         TabAcrossSplits =   -1  'True
         TabAction       =   2
         WrapCellPointer =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1031
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1031
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   7,984
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   7,984
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox pbBottom 
      BackColor       =   &H00C0C0C0&
      Height          =   5052
      Left            =   0
      MousePointer    =   1  'Pfeil
      ScaleHeight     =   417
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1447
      TabIndex        =   0
      Top             =   6240
      Visible         =   0   'False
      Width           =   17412
      Begin MSComctlLib.ProgressBar ProgBar 
         Height          =   252
         Left            =   5640
         TabIndex        =   18
         Top             =   3960
         Width           =   5772
         _ExtentX        =   10181
         _ExtentY        =   445
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.VScrollBar vsbSlide 
         Height          =   4692
         Left            =   17100
         TabIndex        =   17
         Top             =   0
         Width           =   252
      End
      Begin VB.PictureBox picFrame 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'Kein
         Height          =   3840
         Left            =   240
         ScaleHeight     =   320
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   247
         TabIndex        =   13
         Top             =   600
         Width           =   2964
         Begin VB.OptionButton optThumb 
            Height          =   1800
            Index           =   0
            Left            =   0
            Style           =   1  'Grafisch
            TabIndex        =   15
            Top             =   360
            Width           =   1560
         End
         Begin LblCtlsLibUCtl.WindowedLabel Ulabel 
            Height          =   372
            Index           =   0
            Left            =   120
            TabIndex        =   16
            Top             =   0
            Width           =   1200
            _cx             =   2117
            _cy             =   656
            Appearance      =   0
            AutoSize        =   0
            BackColor       =   -2147483633
            BackStyle       =   1
            BorderStyle     =   0
            ClipLastLine    =   -1  'True
            DisabledEvents  =   4099
            DontRedraw      =   0   'False
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483630
            HAlignment      =   0
            HoverTime       =   -1
            MousePointer    =   0
            OwnerDrawn      =   0   'False
            RegisterForOLEDragDrop=   0   'False
            RightToLeft     =   0
            SupportOLEDragImages=   -1  'True
            TextTruncationStyle=   0
            UseMnemonic     =   -1  'True
            UseSystemFont   =   -1  'True
            Text            =   "frmGridAndThumb.frx":04D4
         End
      End
      Begin VB.PictureBox picLoad 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   1560
         Left            =   5640
         ScaleHeight     =   130
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.PictureBox picThumb 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   1200
         Left            =   5640
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   80
         TabIndex        =   1
         Top             =   2280
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Timer TimerfrmGridAndThumb 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   7440
         Top             =   600
      End
      Begin StatBarLibUCtl.StatusBar StatBarU 
         Height          =   348
         Left            =   0
         Top             =   4680
         Width           =   17412
         Version         =   258
         _cx             =   30713
         _cy             =   614
         Appearance      =   0
         BackColor       =   -1
         BorderStyle     =   0
         CustomCapsLockText=   "frmGridAndThumb.frx":0500
         CustomInsertKeyText=   "frmGridAndThumb.frx":0528
         CustomKanaLockText=   "frmGridAndThumb.frx":054E
         CustomNumLockText=   "frmGridAndThumb.frx":0576
         CustomScrollLockText=   "frmGridAndThumb.frx":059C
         DisabledEvents  =   7
         DontRedraw      =   0   'False
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForceSizeGripperDisplay=   0   'False
         HoverTime       =   -1
         MinimumHeight   =   0
         MousePointer    =   0
         BeginProperty Panels {CCA75315-B100-4B5F-80F6-8DFE616F8FDB} 
            Version         =   257
            NumPanels       =   3
            BeginProperty Panel1 {CB0F173F-9E1F-4365-BF3C-6CC52F8C268B} 
               Version         =   258
               Alignment       =   0
               BorderStyle     =   0
               Content         =   0
               Enabled         =   -1  'True
               ForeColor       =   -1
               MinimumWidth    =   100
               PanelData       =   0
               ParseTabs       =   -1  'True
               PreferredWidth  =   150
               RightToLeftText =   0   'False
               Text            =   "frmGridAndThumb.frx":05C4
               Object.ToolTipText     =   "frmGridAndThumb.frx":05EE
            EndProperty
            BeginProperty Panel2 {CB0F173F-9E1F-4365-BF3C-6CC52F8C268B} 
               Version         =   258
               Alignment       =   0
               BorderStyle     =   0
               Content         =   0
               Enabled         =   -1  'True
               ForeColor       =   -1
               MinimumWidth    =   100
               PanelData       =   0
               ParseTabs       =   -1  'True
               PreferredWidth  =   100
               RightToLeftText =   0   'False
               Text            =   "frmGridAndThumb.frx":061C
               Object.ToolTipText     =   "frmGridAndThumb.frx":063C
            EndProperty
            BeginProperty Panel3 {CB0F173F-9E1F-4365-BF3C-6CC52F8C268B} 
               Version         =   258
               Alignment       =   0
               BorderStyle     =   0
               Content         =   0
               Enabled         =   0   'False
               ForeColor       =   -1
               MinimumWidth    =   100
               PanelData       =   0
               ParseTabs       =   -1  'True
               PreferredWidth  =   70
               RightToLeftText =   0   'False
               Text            =   "frmGridAndThumb.frx":066A
               Object.ToolTipText     =   "frmGridAndThumb.frx":0692
            EndProperty
         EndProperty
         PanelToAutoSize =   0
         RegisterForOLEDragDrop=   0   'False
         RightToLeftLayout=   0   'False
         ShowToolTips    =   -1  'True
         SimpleMode      =   0   'False
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   -1  'True
         BeginProperty SimplePanel {CB0F173F-9E1F-4365-BF3C-6CC52F8C268B} 
            Version         =   258
            BorderStyle     =   1
            PanelData       =   0
            ParseTabs       =   -1  'True
            RightToLeftText =   0   'False
            Text            =   "frmGridAndThumb.frx":06C0
            Object.ToolTipText     =   "frmGridAndThumb.frx":06EA
         EndProperty
      End
   End
End
Attribute VB_Name = "frmGridAndThumb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Im Design-Mode eingestellt
'Adodc1.Recordsource = "Fotos"
'Adodc1.CursorLocation = adUseClient
'Adodc1.Cursortype = adOpenStatic
'Adodc1.LockType = adLockOptimistic
'Adodc1.Mode = adModeUnknown
'DbgridNeu.AllowUpdate = True

'Bedingungen für zwei Panele durch horizontalen Splitter geteilt
'Set the form's MousePointer property to "7 - Size NS" so that an up/down sizing cursor will display when the mouse is over the splitter.
'Set each Picture Box control's MousePointer property to "1 - Arrow."
'The program also needs constants for the splitter height and for the minimum permitted height of each pane.

'Der Optionbutton optThumb kann zwar Bilder aus Unicode-Verzeichnissen darstellen, aber seine Caption-Eigenschaft
'kann keine Unicode-Dateinamen darstellen
'Deshalb lege ich über den Optionbutton ein Unicode-Label Ulabel
'Alle für das Erzeugen der Thumbnails gebrauchten Pictureboxen müssen ScaleMode=3 (Pixel) haben

Option Explicit

    Const SPLITTER_HEIGHT = 4
    'Const MIN_PANE_HEIGHT = 80
    Const MIN_PANE_HEIGHT = 1
    
    ' The percentage of the window height
    ' occupied by the top pane.
    Dim TopPanePercent As Single
    ' True when the splitter is being dragged.
    Private Dragging As Boolean


    Dim SQL As String
    Dim msg As String
    Dim blnMarkiert As Boolean
    Public rsDataGrid As ADODB.Recordset
    Dim ZuvorMsg As String
    Dim Bookmark As Variant                                     'Gerbing 20.12.2010
    Dim lngGewählteSpalte As Long                               'Gerbing 01.01.2014
    Public SQLneuHeadClick As String


    Private Type POINTAPI
        x  As Long
        y  As Long
    End Type
    
    Private mbActive                As Boolean
    Private mlCurThumb              As Long
    Private Const SRCCOPY           As Long = &HCC0020
    Private Const STRETCH_HALFTONE  As Long = &H4&
    Private Const SW_RESTORE        As Long = &H9&
    
    Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndParent As Long) As Long
    
    Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
    Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lpPt As POINTAPI) As Long
    Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
    Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
    Private Declare Function UnrealizeObject Lib "gdi32" (ByVal hObject As Long) As Long
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Dim KollIndex As Long
    Public Koll As New Collection
    Private gblnSubdirectories As Boolean
    Dim FolderPath As String
    Dim blnAbbrechen As Boolean
    Public blnComeFromBtnMitThumbnailsClick As Boolean
    Dim blnCreateThumbnailsRunning As Boolean
    
    Private ttToolTip As clsToolTip
    Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
    
    Private TV() As Long                                                'Array speichert die indices der TV=Thumbnails Visible
    '24.11.2016----------------------------------------------ab hier für Video Thumbnails------------------------------------------------
    'From Windows SDK header file propkey.h:
    Private Const SCID_THUMBNAILSTREAM As String = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9},27"
    'Requires Windows XP SP2 or later:
    Private Declare Function PropVariantToVariant Lib "propsys" ( _
        ByRef PropVar As Any, _
        ByRef Var As Variant) As Long
    
    'Private ShellObject
    Public ShellObject As shell32.Shell                                                 'Gerbing 30.11.2016 Public
    'Private ShellObject As Variant                                                     'Gerbing 30.11.2016
    
    Private Const SCID_PerceivedType As String = "{28636AA6-953D-11D2-B5D6-00C04FD918D0},9"
    Private Const SCID_PropStream As String = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9},27"
    
    Private Enum PERCEIVED
        PERCEIVED_TYPE_FIRST = -3
        PERCEIVED_TYPE_CUSTOM = -3
        PERCEIVED_TYPE_UNSPECIFIED = -2
        PERCEIVED_TYPE_FOLDER = -1
        PERCEIVED_TYPE_UNKNOWN = 0
        PERCEIVED_TYPE_TEXT = 1
        PERCEIVED_TYPE_IMAGE = 2
        PERCEIVED_TYPE_AUDIO = 3
        PERCEIVED_TYPE_VIDEO = 4
        PERCEIVED_TYPE_COMPRESSED = 5
        PERCEIVED_TYPE_DOCUMENT = 6
        PERCEIVED_TYPE_SYSTEM = 7
        PERCEIVED_TYPE_APPLICATION = 8
        PERCEIVED_TYPE_GAMEMEDIA = 9
        PERCEIVED_TYPE_CONTACTS = 10
        PERCEIVED_TYPE_LAST = 10
    End Enum
    Private GdipTool As GdipTool
    Dim blnComeFromSlideChange As Boolean
    Public folder As shell32.folder                                     'Gerbing 30.11.2016 Public
'    Private Const VK_SHIFT As Long = &H10&                              'Gerbing 11.04.2017
'    Private Const KeyPressed As Integer = -32767                        'Gerbing 11.04.2017

Private Sub ChangePaneSizes()
    ' Arrange the panes according to the new splitter position.
    Dim TopHeight As Single
    Dim BottomHeight As Single

    ' Do nothing if window is minimized.
    If WindowState = vbMinimized Then Exit Sub
    TopHeight = (ScaleHeight - SPLITTER_HEIGHT) * TopPanePercent
    If TopHeight < MIN_PANE_HEIGHT Then
        TopHeight = MIN_PANE_HEIGHT
    End If
    
    BottomHeight = (ScaleHeight - SPLITTER_HEIGHT) * (1 - TopPanePercent)
    If BottomHeight < MIN_PANE_HEIGHT Then
        BottomHeight = MIN_PANE_HEIGHT
        TopHeight = ScaleHeight - SPLITTER_HEIGHT - BottomHeight
        If TopHeight < MIN_PANE_HEIGHT Then
            TopHeight = MIN_PANE_HEIGHT
        End If
    End If
    pbTop.Move 0, 0, ScaleWidth, TopHeight
    pbBottom.Move 0, TopHeight + SPLITTER_HEIGHT, ScaleWidth, BottomHeight
End Sub


Private Sub btnThumbnailsAbbrechen_Click()
    Dim lIdx As Long
    
    btnThumbnailsAbbrechen.Enabled = False
    btnMitThumbnails.Enabled = True
    btnOhneThumbnails.Enabled = True
    blnAbbrechen = True
    blnCreateThumbnailsRunning = False
    On Error Resume Next
    If optThumb.Count > 2 Then
        For lIdx = 1 To optThumb.Count - 1
            Unload optThumb(lIdx)
            Unload Ulabel(lIdx)
        Next lIdx
    End If
    Screen.MousePointer = vbDefault                                 'Gerbing 22.01.2017
    Call btnOhneThumbnails_Click
End Sub

Private Sub DBGridNeu_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)      'Gerbing 29.10.2019
    If Button = vbRightButton Then
        Call Form1.Hilfebox                                                                         'Gerbing 29.10.2019
    End If
End Sub

Private Sub Form_Load()
    Dim strTemp As String
    
    ' Initially each pane gets half the window.
    ReDim TV(0)
    TopPanePercent = 0.5
    Me.MousePointer = vbSizeNS                                      '7=Size N S (double arrow pointing north and south).
    pbTop.MousePointer = vbArrow                                    '1=Arrow
    pbBottom.MousePointer = vbArrow                                 '1=Arrow
    btnThumbnailsAbbrechen.Enabled = False

    Call AnpassenNutzerWunsch(Me)                       'Gerbing 11.03.2017
    Call AnpassenHeadFont(frmGridAndThumb.DBGridNeu)                'Gerbing 23.06.2011
    If Query.chkFensterGrößeÄnderbar.Value = 1 Then                 'Gerbing 06.12.2005
        Me.Top = Form1.Top
        Me.Left = Form1.Left
        Me.width = Form1.width
    Else
        frmGridAndThumb.Top = 0                                     'Gerbing 16.09.2006
        frmGridAndThumb.Left = 0
        frmGridAndThumb.width = Screen.width
    End If
    'für Laptops mit 1366x768 Pixel Me.Height auf 768-50 Pixel setzen,  'Gerbing 10.02.2017
    'sonst kann man die Unterkante der Form nicht anfassen
    If screenHeight * Screen.TwipsPerPixelY < Me.height Then
        Me.height = 718 * Screen.TwipsPerPixelY
    End If
    DBGridNeu.width = frmGridAndThumb.pbTop.width - 20
    If gblnF5Alt = True Then                                        'Gerbing 22.04.2014
        DBGridNeu.height = frmGridAndThumb.pbTop.height - 40
    Else
        DBGridNeu.height = frmGridAndThumb.pbTop.height - 120       'Gerbing 05.12.2010 29.03.2015
    End If
    DBGridNeu.RowHeight = 250 \ Screen.TwipsPerPixelY               'Gerbing 29.03.2012 29.03.2015
    DBGridNeu.AllowRowSizing = False
    '------------------------------

    'Adodc1.Connect = "Access 2000;"
    btnMerkerspalteEinschalten.tooltipText = LoadResString(2504 + Sprache) 'Mit der Merkerspalte können Sie Fotos vormerken, die Sie exportieren oder löschen oder weiterselektieren wollen
    btnSpaltenbreitenSpeichern.tooltipText = LoadResString(2505 + Sprache) 'Sie können die Spaltenbreite mit der Maus durch Ziehen verändern. Diese Einstellung wird hiermit gespeichert.
    btnRefresh.tooltipText = LoadResString(1126 + Sprache)                  'Aktualisieren-nur sinnvoll in Multiuser-Umgebung       'Gerbing 20.12.2010
    btnSpaltenbreitenSpeichern.Caption = LoadResString(3003 + Sprache) '&Spaltenbreiten speichern
    btnMerkerspalteEinschalten.Caption = LoadResString(3004 + Sprache)  '&Merkerspalte in jedem Datensatz ein/ausschalten
    btnMitThumbnails.Caption = LoadResString(3091 + Sprache) 'mit Thumbnails                                            'Gerbing 03.10.2016
    btnMitThumbnails.tooltipText = LoadResString(2554 + Sprache)   'Je mehr Fotos Sie mit Thumbnails anzeigen, desto spürbarer wird die Verlangsamung
    btnOhneThumbnails.Caption = LoadResString(3092 + Sprache) 'ohne Thumbnails                                          'Gerbing 03.10.2016
    btnOhneThumbnails.tooltipText = LoadResString(2555 + Sprache)  'Das Programm arbeitet schneller ohne Thumbnails
    btnThumbnailsAbbrechen.Caption = LoadResString(3093 + Sprache) 'Thumbnails abbrechen                                'Gerbing 03.10.2016
    
    'Set ShellObject = New shell32.Shell                        'Gerbing 30.11.2016
    On Error Resume Next
    'Set ShellObject = CreateObject("Shell.Application")
    Set ShellObject = CreateObject(CVar("Shell.Application"))
    Set GdipTool = New GdipTool                                 'Gerbing 24.11.2016
    'Subclass the datagrid control to allow mouse wheel usage                       'Gerbing 15.11.2019
   lpPrevWndProc = SetWindowLong(DBGridNeu.hWnd, GWL_WNDPROC, AddressOf WndProc)
   'DBGridNeu.MarqueeStyle = dbgHighlightCell                                        'Gerbing 15.11.2019 default is "Floating Editor"
   DBGridNeu.MarqueeStyle = dbgSolidCellBorder                  'Gerbing 16.11.2019 default is "Floating Editor"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Start dragging the splitter.
    Dragging = True
    If Button = vbRightButton Then
        Call Form1.Hilfebox                                                                                     'Gerbing 29.10.2019
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' As the splitter is dragged.
    ' Do nothing if we're not dragging.
    If Not Dragging Then
        Exit Sub
    End If
    
    TopPanePercent = y / ScaleHeight
    If TopPanePercent < 0 Then TopPanePercent = 0
    If TopPanePercent > 1 Then TopPanePercent = 1
    Call ChangePaneSizes
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' End dragging the splitter.
    Dragging = False
    Call Form_Resize
End Sub

Private Sub Form_Paint()
    Dim msg As String

    On Error Resume Next                                            'Gerbing 15.11.2012
    'frmGridAndThumb.Caption = "  Zur Vollbildansicht eines Bildes Doppel-klicken Sie in die gewünschte Zeile"
    frmGridAndThumb.Caption = LoadResString(1010 + Sprache)              'Gerbing 08.11.2005
    On Error Resume Next
    If Query.CheckDifferenzen.Value = 0 Then                        'Gerbing 09.02.2005
        'frmGridAndThumb.Caption = "Bildanzahl=" & Query.RecordCount & frmGridAndThumb.Caption 'Gerbing 16.06.2005
        frmGridAndThumb.Caption = LoadResString(1011 + Sprache) & Query.RecordCount & frmGridAndThumb.Caption 'Gerbing 08.11.2005
    End If
    If Err = 91 Then    'Objektvariable oder With-Blockvariable nicht festgelegt
        msg = "Es wurde kein einziger Datensatz gefunden." & NL
        msg = msg & "Mit der F8-Taste können Sie die Suche wiederholen"
        'MsgBox msg                                                 'Gerbing 08.11.2005
        MsgBox LoadResString(2007 + Sprache) & NL & LoadResString(2008 + Sprache)
        Exit Sub
    End If
End Sub

Private Sub Form_Resize()
    ' Change pane sizes if the window is resized.
    ChangePaneSizes
    
    pbBottomLeer.Move pbBottom.Left, pbBottom.Top, pbBottom.ScaleWidth + 20, pbBottom.ScaleHeight + 20
    'lblKeineThumbnailsAusgewählt.Caption = "es sind keine Thumbnails ausgewählt oder es sind keine vorhanden"
    lblKeineThumbnailsAusgewählt.Caption = LoadResString(1843 + Sprache)
    lblKeineThumbnailsAusgewählt.Move pbBottomLeer.ScaleWidth \ 2 - lblKeineThumbnailsAusgewählt.width \ 2, _
                                        pbBottom.ScaleHeight \ 2 - lblKeineThumbnailsAusgewählt.height \ 2, _
                                        lblKeineThumbnailsAusgewählt.width, lblKeineThumbnailsAusgewählt.height
    DBGridNeu.width = frmGridAndThumb.pbTop.width - 20              'Gerbing 29.03.2015
    On Error Resume Next
    If gblnF5Alt = True Then                                        'Gerbing 22.04.2014
        DBGridNeu.height = frmGridAndThumb.pbTop.height - 40
    Else
        DBGridNeu.height = frmGridAndThumb.pbTop.height - 120       'Gerbing 05.12.2010 29.03.2015
    End If
    On Error GoTo 0
    If blnComeFromBtnMitThumbnailsClick = True Then
        If glngGridline <> 0 Then
            'Call ChangePicFrameSize(glngGridline)
            Call ChangePicFrameSize(0)
        End If
    End If
End Sub


Private Sub btnMerkerspalteEinschalten_Click()
    'wechselweise Merkerspalte ein/ausschalten
    Dim SQLalt As String
    Dim SQL As String
    Dim SQLMitte As String
    Dim pos As Long
    Dim pos1 As Long
    
    If gblnSchreibgeschützt = True Then                                 'Gerbing 15.05.2006
        msg = gstrFotosMdbLocation & "\Fotos.mdb" & vbNewLine
        'Msg= msg & "Die Datenbank ist schreibgeschützt, Änderungen sind nicht möglich"
        msg = msg & LoadResString(2210 + Sprache)
        'MsgBox Msg
        MessageBoxW 0, StrPtr(msg), StrPtr(LoadResString(1119 + Sprache)), vbInformation
        Exit Sub
    End If
    Me.MousePointer = vbHourglass                                                           'Gerbing 29.07.2007
    '--------------------------------------------------------------------------------------
    'zuerst alle Merkerspalten in der ganzen Tabelle Fotos ausschalten  'Gerbing 26.07.2006
    'SQL = "UPDATE Fotos SET Fotos.Merker = 0;"
    SQL = "UPDATE Fotos SET Fotos." & LoadResString(2524 + Sprache) & " = 0;"
    If gblnSchreibgeschützt = True Then
        ' Recordset erstellen und öffnen adOpenStatic
        Set frmGridAndThumb.rsDataGrid = New ADODB.Recordset
        With frmGridAndThumb.rsDataGrid
            .Source = SQL
            .ActiveConnection = DBado                                                       'Gerbing 23.11.2017
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
    Else
        ' Recordset erstellen und öffnen adOpenDynamic
        Set frmGridAndThumb.rsDataGrid = New ADODB.Recordset
        With frmGridAndThumb.rsDataGrid
            .Source = SQL
            .ActiveConnection = DBado                                                       'Gerbing 23.11.2017
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
    End If
    'SQLalt = Adodc1.RecordSource                                        'Gerbing 24.12.2005
    SQLalt = Query.SQL                                                  'Gerbing 28.08.2006
    pos = InStr(1, SQLalt, "WHERE", vbTextCompare)                      'Gerbing 26.07.2006
    pos1 = InStr(1, SQLalt, "ORDER BY", vbTextCompare)
    SQLMitte = Mid(SQLalt, pos, pos1 - pos) & ";"
    '--------------------------------------------------------------------------------------
    'dann die Merkerspalten bezüglich der Suchkriterien wieder einschalten
    If blnMarkiert = True Then
        'SQL = "UPDATE Fotos SET Fotos.Merker = 1;"
        SQL = "UPDATE Fotos SET Fotos." & LoadResString(2524 + Sprache) & " = 1 "
        SQL = SQL & SQLMitte                                            'Gerbing 26.07.2006
        blnMarkiert = False
    Else
        'SQL = "UPDATE Fotos SET Fotos.Merker = 0;"
        SQL = "UPDATE Fotos SET Fotos." & LoadResString(2524 + Sprache) & " = 0 "
        SQL = SQL & SQLMitte                                            'Gerbing 26.07.2006
        blnMarkiert = True
    End If

    On Error Resume Next
    frmGridAndThumb.rsDataGrid.Close
    On Error GoTo 0
    If gblnSchreibgeschützt = True Then
        ' Recordset erstellen und öffnen adOpenStatic
        Set frmGridAndThumb.rsDataGrid = New ADODB.Recordset
        With frmGridAndThumb.rsDataGrid
            .Source = SQL
            .ActiveConnection = DBado                                                       'Gerbing 23.11.2017
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
    Else
        ' Recordset erstellen und öffnen adOpenDynamic
        Set frmGridAndThumb.rsDataGrid = New ADODB.Recordset
        With frmGridAndThumb.rsDataGrid
            .Source = SQL
            .ActiveConnection = DBado                                                       'Gerbing 23.11.2017
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
    End If
    '------------------------------------------------------
    On Error Resume Next
    frmGridAndThumb.rsDataGrid.Close
    On Error GoTo 0
    If gblnSchreibgeschützt = True Then
        ' Recordset erstellen und öffnen adOpenStatic
        Set frmGridAndThumb.rsDataGrid = New ADODB.Recordset
        With frmGridAndThumb.rsDataGrid
            .Source = SQLalt
            .ActiveConnection = DBado                                                       'Gerbing 23.11.2017
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
    Else
        ' Recordset erstellen und öffnen adOpenDynamic
        Set frmGridAndThumb.rsDataGrid = New ADODB.Recordset
        With frmGridAndThumb.rsDataGrid
            .Source = SQLalt
            .ActiveConnection = DBado                                                       'Gerbing 23.11.2017
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
    End If
    
    Set frmGridAndThumb.Adodc1.Recordset = frmGridAndThumb.rsDataGrid
    Set frmGridAndThumb.DBGridNeu.DataSource = frmGridAndThumb.rsDataGrid
    frmGridAndThumb.DBGridNeu.ReBind

    Call SetSpaltenBreite                               'Gerbing 24.12.2005
    Me.MousePointer = vbNormal                                                              'Gerbing 29.07.2007
End Sub

Private Sub btnShowUsers_Click()
    'Dim cn As New ADODB.Connection
    Dim cn As ADODB.Connection
    Dim Rs As New ADODB.Recordset
    Dim i, j As Long

    Set cn = CreateObject("ADODB Connection")
    
    'das steht auch in der Datei fotos.ldb
    'diese gibts nur in einer Multiuser-Umgebung und wird von selbst gelöscht, wenn es nur noch einen Nutzer gibt
    'cn.Provider = "Microsoft.Jet.OLEDB.4.0"
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"                                                'Gerbing 23.11.2017
    cn.Open "Data Source=" & gstrFotosMdbLocation & "\fotos.mdb"
    ' The user roster is exposed as a provider-specific schema rowset
    ' in the Jet 4 OLE DB provider.  You have to use a GUID to
    ' reference the schema, as provider-specific schemas are not
    ' listed in ADO's type library for schema rowsets

    Set Rs = cn.OpenSchema(adSchemaProviderSpecific, _
    , "{947bb102-5d43-11d1-bdbf-00c04fb92675}")

    'Output the list of all users in the current database.

    Debug.Print Rs.Fields(0).Name, "", Rs.Fields(1).Name, _
    "", Rs.Fields(2).Name, Rs.Fields(3).Name

    While Not Rs.EOF
        Debug.Print Rs.Fields(0), Rs.Fields(1), _
        Rs.Fields(2), Rs.Fields(3)
        Rs.MoveNext
    Wend
End Sub

Private Sub btnSpaltenbreitenSpeichern_Click()
    If gblnSchreibgeschützt = True Then                                 'Gerbing 15.05.2006
        msg = gstrFotosMdbLocation & "\Fotos.mdb" & vbNewLine
        'Msg= msg & "Die Datenbank ist schreibgeschützt, Änderungen sind nicht möglich"
        msg = msg & LoadResString(2210 + Sprache)
        'MsgBox Msg
        MessageBoxW 0, StrPtr(msg), StrPtr(LoadResString(1119 + Sprache)), vbInformation
        Exit Sub
    End If
    Call SpeichernSpaltenBreite
End Sub

Public Sub btnRefresh_Click()
    Dim intLeftcol As Integer
    
    ExportForm.blnExportGestartet = True                                        'Gerbing 26.01.2015 damit nicht kurz das erste Bild aufflackert
    Me.MousePointer = vbHourglass
    intLeftcol = DBGridNeu.LeftCol
    Bookmark = DBGridNeu.Bookmark
    If gblnSchreibgeschützt = True Then
        ' Recordset erstellen und öffnen adOpenStatic
        Set frmGridAndThumb.rsDataGrid = New ADODB.Recordset
        With frmGridAndThumb.rsDataGrid
            .Source = Adodc1.RecordSource
            .ActiveConnection = DBado                                                       'Gerbing 23.11.2017
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
    Else
        ' Recordset erstellen und öffnen adOpenDynamic
        Set frmGridAndThumb.rsDataGrid = New ADODB.Recordset
        With frmGridAndThumb.rsDataGrid
            .Source = Adodc1.RecordSource
            .ActiveConnection = DBado                                                       'Gerbing 23.11.2017
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
    End If
    'Wenn durch Refresh kein einziger Datensatz mehr gefunden wird                          'Gerbing 11.04.2016
    If frmGridAndThumb.rsDataGrid.RecordCount = 0 Then                                      'Gerbing 11.04.2016
        'msg = "Mit diesen Such-Kriterien wurde kein einziger "
        msg = LoadResString(2179 + Sprache)
        'msg = msg & "Datensatz gefunden." & NL
        msg = msg & LoadResString(2180 + Sprache) & NL
        'msg = msg & "Wiederholen Sie die Suche mit anderen Such-Kriterien"
        msg = msg & LoadResString(2181 + Sprache)
        MsgBox msg
        Exit Sub
    End If
    Set frmGridAndThumb.Adodc1.Recordset = frmGridAndThumb.rsDataGrid
    Set frmGridAndThumb.DBGridNeu.DataSource = frmGridAndThumb.rsDataGrid
    frmGridAndThumb.DBGridNeu.ReBind
    On Error Resume Next                                                                    'Gerbing 29.12.2015
    DBGridNeu.Bookmark = Bookmark
    On Error GoTo 0
    'damit wird die komplette Zeile schwarz markiert                                        'Gerbing 30.11.2012
    If frmGridAndThumb.DBGridNeu.SelBookmarks.Count = 1 Then                                'Gerbing 30.11.2012
        frmGridAndThumb.DBGridNeu.SelBookmarks.Remove 0                                     'Gerbing 30.11.2012
    End If                                                                                  'Gerbing 30.11.2012
    frmGridAndThumb.DBGridNeu.SelBookmarks.Add frmGridAndThumb.rsDataGrid.Bookmark          'Gerbing 30.11.2012
    Call SetSpaltenBreite
    'Horizontalen Scrollbalken wieder so einstellen wie vor dem Sortieren
    On Error Resume Next
    DBGridNeu.Scroll intLeftcol, 0
    On Error GoTo 0
    Me.MousePointer = vbDefault
    ExportForm.blnExportGestartet = False                                                   'Gerbing 26.01.2015
End Sub

Private Sub btnMitThumbnails_Click()
    Dim msg As String
    Dim lIdx As Long
    Dim SmallChange As Integer                                                  'Gerbing 10.11.2016
    Dim LargeChange As Integer
    
    vsbSlide.Visible = False                                                    'Gerbing 10.11.2016
    btnOhneThumbnails.Enabled = False
    If optThumb.Count > 2 Then
        On Error Resume Next
        For lIdx = 1 To optThumb.Count - 1
            Unload optThumb(lIdx)
            Unload Ulabel(lIdx)
        Next lIdx
        On Error GoTo 0
    End If
    TopPanePercent = 0.5
    Call ChangePaneSizes
    blnComeFromBtnMitThumbnailsClick = True

    pbBottom.Visible = True
    pbBottomLeer.Visible = False
    Set Koll = Nothing
    Call KollFüllen
    If Koll.Count >= 32766 Then                                                 'Gerbing 10.11.2016
        MsgBox "Thumbnails Maximum = 32766"
        Set Koll = Nothing
    End If
    If Koll.Count <> 0 Then
        btnThumbnailsAbbrechen.Enabled = True
        btnMitThumbnails.Enabled = False
        btnOhneThumbnails.Enabled = False
        SetParent ProgBar.hWnd, StatBarU.hWnd
        StatBarU.Panels(0).PreferredWidth = (Me.width \ Screen.TwipsPerPixelX) \ 4
        StatBarU.Panels(1).PreferredWidth = (Me.width \ Screen.TwipsPerPixelX) \ 4
        StatBarU.Panels(2).PreferredWidth = (Me.width \ Screen.TwipsPerPixelX) \ 4
        StatBarU_ResizedControlWindow
'        Progbar.Max = Koll.Count
        blnCreateThumbnailsRunning = True
        
        Call CreateThumbs
        
        If blnAbbrechen = False Then
            vsbSlide.Value = 1                                                      'Gerbing 10.11.2016
            If Koll.Count > 0 Then
                vsbSlide.Max = optThumb.Count
                vsbSlide.Min = 1
                SmallChange = vsbSlide.Max \ 100
                LargeChange = vsbSlide.Max \ 10
                If SmallChange > 0 Then
                    vsbSlide.SmallChange = SmallChange
                End If
                If LargeChange > 0 Then
                    vsbSlide.LargeChange = LargeChange
                End If
                vsbSlide.Visible = True
            End If
        End If
        blnCreateThumbnailsRunning = False
        btnThumbnailsAbbrechen.Enabled = False
        btnMitThumbnails.Enabled = True
        btnOhneThumbnails.Enabled = True
    Else
        Call btnOhneThumbnails_Click
        btnThumbnailsAbbrechen.Enabled = False
        btnMitThumbnails.Enabled = True
        btnOhneThumbnails.Enabled = True
        pbBottom.Visible = False
        pbBottomLeer.Visible = True
'        Msg = "Die Datenbank enthält keine Bilder, oder" & vbNewLine
'        Msg = Msg & "es konnten keine Thumbnails erzeugt werden."
        msg = LoadResString(2337 + Sprache) & vbNewLine     'Die Datenbank enthält keine Bilder, oder
        msg = msg & LoadResString(2338 + Sprache)           'es konnten keine Thumbnails erzeugt werden.
        MsgBox msg
    End If
End Sub

Private Sub btnOhneThumbnails_Click()
    Dim lIdx As Long
            
    blnComeFromBtnMitThumbnailsClick = False
    btnThumbnailsAbbrechen.Enabled = False
    Set Koll = Nothing
    
    'TopPanePercent = 0.95
    TopPanePercent = 0.999
    Call ChangePaneSizes
    If optThumb.Count > 2 Then
        For lIdx = 1 To optThumb.Count - 1
            On Error Resume Next
            Unload optThumb(lIdx)
            Unload Ulabel(lIdx)
            If Err.Number <> 0 Then Exit For
        Next lIdx
        On Error GoTo 0
    End If
    pbBottom.Visible = False
    pbBottomLeer.Visible = True
    glngGridline = 0
    Call Form_Resize
End Sub

Private Sub DBGridNeu_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Dim Merker                  'Gerbing 02.05.2006
    Dim Jahr As String
    Dim n As Long
    Dim strTemp As String
    Dim maximum As String
    Dim pos1 As Long
    'Gerbing 20.04.2020 Anstelle von zB If ColIndex = 6 Then benutze
    'zB If DBGridNeu.Columns(ColIndex).DataField = LoadResString(1028 + Sprache) 1028=Dateiname

    If DBGridNeu.Columns(ColIndex).DataField = "GPSLongitude" Or DBGridNeu.Columns(ColIndex).DataField = "GPSLatitude" Then 'Gerbing 02.10.2019
        'MsgBox "Manuelle Änderungen im Feld GPSLongitude oder GPSLatitude sind verboten. Benutzen Sie im Menü Datei... -> Einfügen Geo-Position"
        MsgBox LoadResString(1155 + Sprache)
        Cancel = True
        Exit Sub
    End If
    If DBGridNeu.Columns(ColIndex).DataField = LoadResString(1028 + Sprache) Then           '1028=Dateiname Gerbing 20.04.2020
        'Msg = "Änderungen in diesem Feld sind verboten." & vbNewLine
        'Msg = Msg & "Benutzen Sie RenamMdb zur übereinstimmenden Änderung in der Datenbank und im Ordner."
        MsgBox LoadResString(2002 + Sprache) & vbNewLine & LoadResString(2003 + Sprache)
        Cancel = True
        Exit Sub
    End If                                                                                  'Gerbing 20.04.2020
    If DBGridNeu.Columns(ColIndex).DataField = LoadResString(1031 + Sprache) Then           '1031=DateinameKurz Gerbing 20.04.2020
        'Msg = "Änderungen in diesem Feld sind verboten." & vbNewLine
        'Msg = Msg & "Benutzen Sie RenamMdb zur übereinstimmenden Änderung in der Datenbank und im Ordner."
        MsgBox LoadResString(2002 + Sprache) & vbNewLine & LoadResString(2003 + Sprache)
        Cancel = True
        Exit Sub
    End If                                                                                  'Gerbing 20.04.2020
    If DBGridNeu.Columns(ColIndex).DataField = LoadResString(1032 + Sprache) Then           '1032=DDatum Gerbing 20.04.2020
        'Msg = "Änderungen in diesem Feld sind verboten." & vbNewLine
        'Msg = Msg & "Benutzen Sie RenamMdb zur übereinstimmenden Änderung in der Datenbank und im Ordner."
        MsgBox LoadResString(2002 + Sprache) & vbNewLine & LoadResString(2003 + Sprache)
        Cancel = True
        Exit Sub
    End If                                                                                  'Gerbing 20.04.2020
    
    
    If DBGridNeu.Columns(ColIndex).DataField = LoadResString(1025 + Sprache) Then           '1025=Ort Gerbing 20.04.2020
    'If ColIndex = 3 Then                                                            'Gerbing 16.10.2014
        'Ort
        'Nur bei meiner privaten Datenbank(nicht SQL) soll die Gültigkeitsprüfung stattfinden
        If Gefundenexifdatetimeoriginal = True And gblnVollversion = True And Sprache = 0 And gblnSQLServerVersion = False Then 'Gerbing 08.03.2016
            strTemp = oCat.Tables("fotos").Columns("ort").Properties(8)                 '8=ValidationRule=Len([Ort])<33
            If strTemp <> "" Then
                'es gibt eine Gültigkeitsregel
                pos1 = InStr(1, strTemp, "<", vbTextCompare)
                If pos1 <> 0 Then
                    maximum = Mid(strTemp, pos1 + 1, Len(strTemp) - pos1)
                    If IsNumeric(maximum) Then
                        Trim (DBGridNeu.Columns(3).Text)
                        If Not Len(DBGridNeu.Columns(3).Text) < maximum Then
                            MsgBox "Gültigkeitsregel: " & oCat.Tables("fotos").Columns("ort").Properties(7)    '7=ValidationText
                            'DBGridNeu.Columns(3).Text = left(DBGridNeu.Columns(3).Text, 32)           'begrenzen auf maximal erlaubte Bytes
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
    End If
    If DBGridNeu.Columns(ColIndex).DataField = LoadResString(1025 + Sprache) Then           '1030=Kommentar Gerbing 20.04.2020
    'If ColIndex = 8 Then                                                            'Gerbing 16.10.2014
        'Kommentar
        If Gefundenexifdatetimeoriginal = True And gblnVollversion = True And Sprache = 0 And gblnSQLServerVersion = False Then 'Gerbing 08.03.2016
            strTemp = oCat.Tables("fotos").Columns("Kommentar").Properties(8)           '8=ValidationRule=Len([Kommentar])<2001
            If strTemp <> "" Then
                'es gibt eine Gültigkeitsregel
                pos1 = InStr(1, strTemp, "<", vbTextCompare)
                If pos1 <> 0 Then
                    maximum = Mid(strTemp, pos1 + 1, Len(strTemp) - pos1)
                    If IsNumeric(maximum) Then
                        Trim (DBGridNeu.Columns(8).Text)
                        If Not Len(DBGridNeu.Columns(8).Text) < maximum Then
                            MsgBox "Gültigkeitsregel: " & oCat.Tables("fotos").Columns("Kommentar").Properties(7)    '7=ValidationText
                            'DBGridNeu.Columns(8).Text = left(DBGridNeu.Columns(8).Text, 2000)           'begrenzen auf maximal erlaubte Bytes
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
    End If
    'Bei Dateinamen mit Hochkomma führt 'rstsql.Fields("IPTCPresent") = 0' zu Laufzeitfehler -> ersetzen durch 2 Hochkommas 'Gerbing 23.01.2018
    strTemp = DBGridNeu.Columns(6).Text
    strTemp = Replace(strTemp, "'", "''")                                                   'Gerbing 23.01.2018
    'IPTCPresent = 0 setzen
    If StrComp(Right(strTemp, 3), "jpg", vbTextCompare) = 0 Then                            'Gerbing 23.01.2018
        'nur wenn die rechten 3 Bytes des Dateinamens = jpg
        Select Case ColIndex                                                                'Gerbing 10.04.2016
            Case 2, 3, 4, 5, 8              '2=Situation 3=Ort 4=Land 5=Personen 8=Kommentar
                'Wenn sich der Inhalt dieser Spalte ändert soll "IPTCPresent" = 0 werden
                If Trim(OldValue) <> Trim(DBGridNeu.Columns(ColIndex).Text) Then
                    'SQL = "Select * FROM Fotos Where Dateiname=" & """" & DBGridNeu.Columns(6).Text  & """"
                    SQL = "Select * FROM Fotos Where " & LoadResString(1028 + Sprache) & "='" & strTemp & "'"   'Gerbing 23.01.2018
                    With rstsql
                        .Source = SQL
                        .ActiveConnection = DBado                                                       'Gerbing 23.11.2017
                        .CursorType = adOpenDynamic
                        .LockType = adLockOptimistic
                        .CursorLocation = adUseClient
                        .Open
                    End With
                    Do Until rstsql.EOF
                        On Error Resume Next            'zur Fehlerabwehr wenn eine Spalte zB auf 2 Zeichen begrenzt ist
                        rstsql.Fields("IPTCPresent") = 0
                        On Error GoTo 0
                        rstsql.Update
                        rstsql.MoveNext
                    Loop
                    rstsql.Close
                End If
        End Select
    End If                                                                                  'Gerbing 10.04.2016
    If DBGridNeu.Columns(ColIndex).DataField = LoadResString(1106 + Sprache) Then           '1106=BreitePixel Gerbing 20.04.2020
    'If ColIndex = 11 Then                                                                   'Gerbing 31.12.2007
        If Trim(OldValue) <> Trim(DBGridNeu.Columns(ColIndex).Text) Then
            'BreitePixel
    '        msg = "Änderungen in diesem Feld sind verboten." & vbNewLine
    '        msg = msg & "Benutzen Sie fotosmdb (Prüfen1), wenn Sie BreitePixel neu berechnen wollen."
            'MsgBox msg                                             'Gerbing 08.11.2005
            MsgBox LoadResString(2002 + Sprache) & vbNewLine & LoadResString(2005 + Sprache)
            Cancel = True
            Exit Sub
        End If
    End If
    '---------------------------------------------------------------------------------------------------------
    If DBGridNeu.Columns(ColIndex).DataField = LoadResString(1107 + Sprache) Then           '1107=HoehePixel Gerbing 20.04.2020
    'If ColIndex = 12 Then                                                                  'Gerbing 31.12.2007
        'If Trim(OldValue) <> Trim(DBGridNeu.Columns(12).Text) Then
        If Trim(OldValue) <> Trim(DBGridNeu.Columns(ColIndex).Text) Then                    'Gerbing 02.07.2020
            'HoehePixel
    '        msg = "Änderungen in diesem Feld sind verboten." & vbNewLine
    '        msg = msg & "Benutzen Sie fotosmdb (Prüfen1), wenn Sie HoehePixel neu berechnen wollen."
            'MsgBox msg                                             'Gerbing 08.11.2005
            MsgBox LoadResString(2002 + Sprache) & vbNewLine & LoadResString(2006 + Sprache)
            Cancel = True
            Exit Sub
        End If
    End If
    '---------------------------------------------------------------------------------------------------------
    If DBGridNeu.Columns(ColIndex).DataField = LoadResString(2537 + Sprache) Then           '2537=AudioFileExists Gerbing 20.04.2020
    'If ColIndex = 13 Then                                                                   'Gerbing 31.12.2007
        If Trim(OldValue) <> Trim(DBGridNeu.Columns(ColIndex).Text) Then
            'AudioFileExists                                                                'Gerbing 31.12.2007
    '        msg = "Änderungen in diesem Feld sind verboten." & vbNewLine
    '        msg = msg & "Benutzen Sie fotosmdb (PrüfenS), wenn Sie Differenzen zwischen Audio-Kommentaren und der Spalte 'AudioFileExist' reparieren wollen"
            'MsgBox msg                                                                     'Gerbing 31.12.2007
            MsgBox LoadResString(2002 + Sprache) & vbNewLine & LoadResString(2265 + Sprache)
            Cancel = True
            Exit Sub
        End If
    End If
    '---------------------------------------------------------------------------------------------------------
    If DBGridNeu.Columns(ColIndex).DataField = "IPTCPresent" Then                           'Gerbing 20.04.2020
    'If ColIndex = 14 Then                                                                   'Gerbing 16.04.2008
        If Trim(OldValue) <> Trim(DBGridNeu.Columns(ColIndex).Text) Then
            'IPTCPresent                                                                    'Gerbing 16.04.2008
    '        msg = "Änderungen in diesem Feld sind verboten." & vbNewLine
    '        msg = msg & "Benutzen Sie fotosmdb (PrüfenIPTC), wenn Sie das Feld IPTCPresent neu berechnen lassen wollen"
            'MsgBox msg                                                                     'Gerbing 31.12.2007
            MsgBox LoadResString(2002 + Sprache) & vbNewLine & LoadResString(3146 + Sprache)
            Cancel = True
            Exit Sub
        End If
    End If
    '-----------------------------------------------------------------------------------------------
    'Kontrolle, ob in Spalte SWF erlaubter Inhalt steht                                     'Gerbing 31.12.2007
    If DBGridNeu.Columns(DBGridNeu.Col).Caption = LoadResString(1029 + Sprache) Then        '1029=SWF
        strTemp = UCase(DBGridNeu.Text)
        If Sprache = 0 Then
            'das ist deutsch
            Select Case strTemp
                Case "SW", "F", "FV", "SV", "BW", "C", "CV", "BV"
                
                Case Else
                    'MsgBox "Falscher Wert in Spalte SWF"
                    MsgBox LoadResString(2264 + Sprache)
                    Cancel = True
                    Exit Sub
            End Select
        Else
            'das ist english
            Select Case strTemp
                Case "SW", "F", "FV", "SV", "BW", "C", "CV", "BV"
                
                Case Else
                    'MsgBox "Falscher Wert in Spalte SWF"
                    MsgBox LoadResString(2264 + Sprache)
                    Cancel = True
                    Exit Sub
            End Select
        End If
    End If
    '-----------------------------------------------------------------------------------------------
    'If DBGridNeu.Col = 0 Then 'Gerbing 21.03.2006                       'DBGridNeu.Col = 0 = Merker
    If DBGridNeu.Columns(ColIndex).DataField = LoadResString(2524 + Sprache) Then       '2524=Merker        'Gerbing 02.07.2020
        Merker = DBGridNeu.Columns(ColIndex).Value                                                          'Gerbing 02.07.2020
        Select Case Merker
            Case "0"

            Case "1"

            Case "-1"

            Case Else
                'MsgBox "Sie dürfen in die Spalte Merker nur 0 oder 1 eintragen"
                MsgBox LoadResString(2001 + Sprache)            'Gerbing 08.11.2005
                DBGridNeu.Columns(ColIndex).Value = "0"                   'Gerbing 21.03.2006               'Gerbing 02.07.2020
                Cancel = True
                Exit Sub
        End Select
    End If
    '-----------------------------------------------------------------------------------------------
    'If DBGridNeu.Col = 1 Then                           'Gerbing 14.05.2006 DBGridNeu.Col = 1 = Jahr
    If DBGridNeu.Columns(ColIndex).DataField = LoadResString(1023 + Sprache) Then       '1023=Jahr          'Gerbing 02.07.2020
        Jahr = DBGridNeu.Columns(ColIndex).Text                'Gerbing 21.03.2006                          'Gerbing 02.07.2020
        If Len(Jahr) <> 4 Then
            'MsgBox "Jahr muß eine 4-stellige Zahl sein"
            MsgBox LoadResString(2127 + Sprache)
            Cancel = True
            Exit Sub
        End If
        If Not IsNumeric(Jahr) Then
            'MsgBox "Jahr muß eine 4-stellige Zahl sein"
            MsgBox LoadResString(2127 + Sprache)
            Cancel = True
            Exit Sub
        End If
        'MsgBox "Klicken Sie nach Änderung des Feldes Jahr in die Nachbarspalte der gleichen Zeile" 'Gerbing 14.05.2006
        MsgBox LoadResString(2249 + Sprache)
    End If
    '-----------------------------------------------------------------------------------------------
    'Falls 'erster Treffer pro Jahr', dann Stichwortänderung in Tabelle Fotos beim aktuellen Dateiname nachziehen
    If Query.optNurErstenTreffer.Value = True Or Query.optErsterZufallstreffer.Value = True Then        'Gerbing 09.02.2013
'        msg = "Spalte=" & DBGridNeu.Columns(DBGridNeu.Col).Caption & vbNewLine
'        msg = msg & "Inhalt=" & DBGridNeu.Text & vbNewLine
'        msg = msg & "Dateiname=" & DBGridNeu.Columns(6).Text
'        MsgBox msg
        'SQL = "Select * FROM Fotos Where Dateiname=" & """" & DBGridNeu.Columns(6).Text  & """"
        'Bei Dateinamen mit Hochkomma kommt Laufzeitfehler -> ersetzen durch 2 Hochkommas
        strTemp = DBGridNeu.Columns(6).Text
        strTemp = Replace(strTemp, "'", "''")                                                       'Gerbing 23.01.2018
        SQL = "Select * FROM Fotos Where " & LoadResString(1028 + Sprache) & "='" & strTemp & "'"   'Gerbing 23.01.2018
        With rstsql
            .Source = SQL
            .ActiveConnection = DBado                                                       'Gerbing 23.11.2017
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
        Do Until rstsql.EOF
            On Error Resume Next            'zur Fehlerabwehr wenn eine Spalte zB auf 2 Zeichen begrenzt ist
            rstsql.Fields(DBGridNeu.Columns(DBGridNeu.Col).Caption) = DBGridNeu.Text
            On Error GoTo 0
            rstsql.Update
            rstsql.MoveNext
        Loop
        rstsql.Close
        Call Query.FrageObNurErstenTreffer                                                  'Gerbing 26.08.2008
        'Hier wird 'Set DBGridNeu.Datasource' erneut ausgeführt weil dem Grid DBGridNeu eine Abfrage mit
        'Inner Join zugrunde liegt und ein solches Recordset cannot be updated
    End If
End Sub

Private Sub DBGridNeu_BeforeUpdate(Cancel As Integer)
    'MsgBox "DBGridNeu_BeforeUpdate"                                                         'Gerbing 15.10.2014
End Sub

Private Sub DbGridNeu_DblClick()
    'es gibt irgendein Problem, wenn ich in der frmGridAndThumb einen Doppelklick auf eine Spaltenüberschrift
    'ausführe. Das Image1 verrutscht außerhalb der Zentrierung. Es sieht so aus, als würde der Doppelklick
    'gleichzeitig als Verschiebeklick aufgefasst.
    'Ich erfinde den Schalter gblnfrmGridAndThumbDblClick            'Gerbing 06.06.2005
    gblnfrmGridAndThumbDblClick = True                               'Gerbing 06.06.2005
    On Error Resume Next                                        'Gerbing 25.06.2015
    frmGridAndThumb.Hide
    Hilfebx.Hide                                                'Gerbing 16.09.2004
    On Error GoTo 0                                             'Gerbing 25.06.2015
    Call Form1.BildAnzeigen
End Sub

Private Sub DBGridNeu_Error(ByVal DataError As Integer, Response As Integer)
    'MsgBox "DbGridNeu_Error " & DBGridNeu.ErrorText

    Response = 0
    On Error Resume Next
    Adodc1.Recordset.CancelUpdate
    On Error GoTo 0
End Sub

Private Sub DbGridNeu_HeadClick(ByVal ColIndex As Integer)
    'In dieser Prozedur funktioniert keinerlei Mousepointer auf Hourglass setzen
    'Diese Prozedur kann nur richtig arbeiten, wenn es im SQL String eine ORDER BY Anweisung gibt
    'es gibt 4 Varianten wie ein SQL String aufgebaut wird. Mit allen 4 Varianten muss diese Prozedur
    'fertig werden.
    '1. Normale Suche nach Suchkriterien
    '2. Suchen Differenzen Jahr
    '3. Gespeicherte Abfrage
    '4. ErsterTreffer
    
    Dim SQL As String
    Dim SQLalt As String
    Dim pos As Long
    Dim Links As String
    Dim ColCaption As String
    Dim intLeftcol As Integer
    Dim n As Long
    
    'DoEvents 'auskommentiert Gerbing 10.09.2009
    gblnWasHeadClick = True                                             'Gerbing 29.03.2015
    intLeftcol = DBGridNeu.LeftCol                                      'Gerbing 13.03.2005
    ColCaption = DBGridNeu.Columns(ColIndex).Caption
    'SQLalt = Query.SQL                                                 'Gerbing 16.06.2005
    SQLalt = Adodc1.RecordSource                                        'Gerbing 24.12.2005
    '----------------------------------------------------------------------------------------
    pos = InStr(1, SQLalt, "DESC", vbTextCompare)
    If pos <> 0 Then
        SQL = " ORDER BY [" & ColCaption & "];"                        'Gerbing 22.02.2005
    Else
        SQL = " ORDER BY [" & ColCaption & "] DESC;"
    End If
    pos = InStr(1, SQLalt, "ORDER BY", vbTextCompare)
    If pos <> 0 Then                                                    'Gerbing 20.06.2006
        Links = Left(SQLalt, pos - 1)
    Else
        Links = SQLalt
        'wenn ein Semikolon am Ende steht, dann abschneiden
        pos = InStr(1, Links, ";")
        If pos <> 0 Then
            Links = Mid(Links, 1, pos - 1)
        End If
    End If
    SQLneuHeadClick = Links & SQL
    'Query.SQL = SQLneuHeadClick                                                  'Gerbing 16.06.2005
    
    On Error Resume Next
    frmGridAndThumb.rsDataGrid.Close
    On Error GoTo 0
    
    If gblnSchreibgeschützt = True Then
        ' Recordset erstellen und öffnen
        Set frmGridAndThumb.rsDataGrid = New ADODB.Recordset
        With frmGridAndThumb.rsDataGrid
            .Source = SQLneuHeadClick
            .ActiveConnection = DBado                                                       'Gerbing 23.11.2017
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
    Else
        ' Recordset erstellen und öffnen
        Set frmGridAndThumb.rsDataGrid = New ADODB.Recordset
        With frmGridAndThumb.rsDataGrid
            .Source = SQLneuHeadClick
            .ActiveConnection = DBado                                                       'Gerbing 23.11.2017
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
    End If
    Set frmGridAndThumb.Adodc1.Recordset = frmGridAndThumb.rsDataGrid
    Set frmGridAndThumb.DBGridNeu.DataSource = frmGridAndThumb.rsDataGrid
    frmGridAndThumb.DBGridNeu.ReBind
    
    Call SetSpaltenBreite
    'Horizontalen Scrollbalken wieder so einstellen wie vor dem Sortieren
    On Error Resume Next
    DBGridNeu.Scroll intLeftcol, 0
    On Error GoTo 0
    frmGridAndThumb.DBGridNeu.AllowUpdate = True
    If blnComeFromBtnMitThumbnailsClick = False Then
        Call btnOhneThumbnails_Click
    Else
        Call btnMitThumbnails_Click
    End If
    'gblnWasHeadClick = True
End Sub

Private Sub DBGridNeu_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyMultiply And Shift = vbCtrlMask Then 'vbKeyMultiply ist das Multiplikationszeichen auf dem Ziffernblock 'Gerbing 01.01.2014 12.02.2014

        lngGewählteSpalte = DBGridNeu.Col
        'DBGridNeu.Text in die Zwischenablage kopieren
        ClipboardSetText Me.hWnd, DBGridNeu.Text                                                'Gerbing 04.02.2014
        KeyCode = 0                             'sonst wird in die Zelle 'c' eingetragen
        Exit Sub
    End If
    '------------------------------------------------------------------------------------------------------
    If KeyCode = 40 Or KeyCode = 38 Then                                                        'Gerbing 04.03.2013
        If frmGridAndThumb.DBGridNeu.SelBookmarks.Count = 1 Then                                'Gerbing 04.03.2013
            frmGridAndThumb.DBGridNeu.SelBookmarks.Remove 0                                     'Gerbing 04.03.2013
        End If                                                                                  'Gerbing 04.03.2013
    End If
    If KeyCode = 40 Then 'Pfeil nach unten
        'damit wird die aktuelle Zeile schwarz
        frmGridAndThumb.DBGridNeu.SelBookmarks.Add frmGridAndThumb.rsDataGrid.Bookmark + 1      'Gerbing 04.03.2013
    End If
    If KeyCode = 38 Then 'Pfeil nach oben
        'damit wird die aktuelle Zeile schwarz
        frmGridAndThumb.DBGridNeu.SelBookmarks.Add frmGridAndThumb.rsDataGrid.Bookmark - 1      'Gerbing 04.03.2013
    End If
    Select Case KeyCode
        Case vbKeyF1, vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF8, vbKeyF10                              'Gerbing 21.09.2014
            'Debug.Print "DBGridNeu_KeyDown-Case vbKeyF1, vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF8, vbKeyF10 "
            Me.Hide
            'Tastatur-Eingabe weiterreichen
            '-> Public Sub Form_    (KeyCode As Integer, Shift As Integer)
                Call Form1.Form_KeyDown(KeyCode, Shift)
        Case vbKeyF5                                                                            'Gerbing 21.09.2014
            If Shift = vbShiftMask Then
                Me.Hide
                'Tastatur-Eingabe weiterreichen
                '-> Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
                    Call Form1.Form_KeyDown(KeyCode, Shift)
            End If
        Case vbKeyG
            If Shift = vbCtrlMask Then                                                          'Gerbing 03.10.2016
                KeyCode = 0                 'sonst wird in die Zelle 'g' eingetragen            'Gerbing 02.07.2020
                Me.Hide
                'Tastatur-Eingabe weiterreichen
                '-> Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
                    Call Form1.Form_KeyDown(KeyCode, Shift)
            End If
    End Select
End Sub

Private Sub DBGridNeu_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)    'Gerbing 01.01.2014
    Dim i As Long
    
    If Shift = vbShiftMask + vbCtrlMask Then
        'Kopieren mit Multiselect
        y = y \ Screen.TwipsPerPixelY                                                           'Gerbing 29.03.2015 weil jetzt Pixel und bisher Twips
        HandleDTGMultiSelect DBGridNeu, y ', LastRow, dtgGDLookup.row
        'If lngGewählteSpalte <> 0 Then                                                         'Gerbing 22.09.2016 auskommentiert sonst wird Merkerspalte nicht bedient
            Select Case lngGewählteSpalte
                Case 6, 9, 10, 11, 12, 13, 14, 15, 16                                           'das sind die verbotenen Spalten
                    'Gerbing 02.10.2019 zusätzlich verbotene Spalten 15=GPSLatitude und 15=GPSLongitude
                    'MsgBox "verbotene Spalte"
                Case -1
                    'da ist keine Spalte gewählt worden                                         'Gerbing 13.04.2014
                Case Else
                    For i = 0 To DBGridNeu.SelBookmarks.Count - 1
                        gblnComeFromF2F3 = True                                                 'Gerbing 27.03.2014
                        DBGridNeu.Col = lngGewählteSpalte
                        DBGridNeu.FirstRow = DBGridNeu.SelBookmarks(i) - 1                      'DBGridNeu.Row muss sichtbar sein, sonst
                        If DBGridNeu.SelBookmarks(i) - DBGridNeu.FirstRow >= 0 Then
                            DBGridNeu.Row = DBGridNeu.SelBookmarks(i) - DBGridNeu.FirstRow      'Laufzeitfehler 6148 Ungültige Zeilennummer
                        End If
                        'DBGridNeu.Text = Clipboard.GetText(1)
                         'DBGridNeu.Text wieder aus der Zwischenablage auslesen
                        DBGridNeu.Text = ClipboardGetText(Me.hWnd)                              'Gerbing 04.02.2014
                    Next i
            End Select
        'End If                                                                                 'Gerbing 22.09.2016 auskommentiert sonst wird Merkerspalte nicht bedient
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1, vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF8, vbKeyF10                              'Gerbing 26.09.2013
            'Debug.Print "Form_KeyDown-Case vbKeyF1, vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF8, vbKeyF10"
            Me.Hide
            'Tastatur-Eingabe weiterreichen
            '-> Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
                Call Form1.Form_KeyDown(KeyCode, Shift)
        Case vbKeyF5                                                                            'Gerbing 21.09.2014
            If Shift = vbShiftMask Then
                Me.Hide
                'Tastatur-Eingabe weiterreichen
                '-> Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
                    Call Form1.Form_KeyDown(KeyCode, Shift)
            End If
        Case vbKeyG
            If Shift = vbCtrlMask Then                                                          'Gerbing 03.10.2016
                Me.Hide
                'Tastatur-Eingabe weiterreichen
                '-> Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
                    Call Form1.Form_KeyDown(KeyCode, Shift)
            End If
    End Select
End Sub

Private Sub DbGridNeu_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    'Bei RowColChange stehen noch die Daten der Zeile zur Verfügung die den Fokus verloren hat
    'Bei SelChange stehen die Daten der Zeile zur Verfügung die den Fokus bekommen hat
    
    Dim DateinamenErweiterung As String
    Dim CellVal As String
    Dim intLänge As Integer
    Dim strTemp As String
    Dim pos As Long                                                                         'Gerbing 29.03.2015
    Dim MerkeIndex As Long
    Dim i As Long
    Dim blnThumbnailGefunden As Boolean                                                     'Gerbing 10.11.2016

    If blnComeFromBtnMitThumbnailsClick = False Then                                        'Gerbing 24.06.2015
        If gblnVollversion = False Then
            If gintDiffTage > 365 Then
                If Sprache = 0 Then
                    MsgBox "Sie benutzen die Shareware-Version von GERBING Fotoalbum."      'Gerbing 25.06.2015
                Else
                    MsgBox "You use the Shareware version of GERBING Fotoalbum."
                End If
                'Copy.Show 1                                                                'Gerbing 24.06.2015
            End If
        End If
    End If
    If Adodc1.Recordset.EOF Or Adodc1.Recordset.BOF Then
        Exit Sub
    End If
    If gblnComeFromThumbs = True Then                                                       'Gerbing 29.03.2015
        Exit Sub
    End If
    On Error Resume Next                            'Gerbing 21.03.2006
    'Falls 'erster Treffer pro Jahr', kein Update weil wegen Inner Join cannot be updated   'Gerbing 26.08.2008
    If Query.optNurErstenTreffer.Value = False Then                                         'Gerbing 26.08.2008
        Adodc1.Recordset.Update                                                             'Gerbing 26.08.2008
    End If                                                                                  'Gerbing 26.08.2008
    'On Error GoTo 0                                 'Gerbing 21.03.2006
    On Error GoTo RowColChangeError                                                         'Gerbing 25.06.2013
    gstrRowColChangeName = Adodc1.Recordset.Fields(LoadResString(1028 + Sprache))
    If Mid(gstrRowColChangeName, Len(gstrRowColChangeName) - 3, 1) = "." Then             'Gerbing 25.06.2006
        intLänge = 3
    End If
    If Mid(gstrRowColChangeName, Len(gstrRowColChangeName) - 4, 1) = "." Then
        intLänge = 4
    End If
    If Mid(gstrRowColChangeName, Len(gstrRowColChangeName) - 5, 1) = "." Then
        intLänge = 5
    End If
    DateinamenErweiterung = Right(gstrRowColChangeName, intLänge)
    DateinamenErweiterung = UCase(DateinamenErweiterung)            'Gerbing 26.11.2006
    
    'Gerbing 26.11.2006
    If ExportForm.blnExportGestartet = False Then                               'Gerbing 08.12.2006
        '--------------------------------------------------------------------------------------
        'Messen der Ausführungsdauer
        'glngStartMillisek = timeGetTime
        If Replace(gstrRowColChangeName, "+:\", gstrFotosMdbLocation & "\") = gstrFRODN Then    'Gerbing 27.03.2014
            gblnComeFromF2F3 = True                                                             'Gerbing 27.03.2014
        End If                                                                                  'Gerbing 27.03.2014
        If gblnComeFromF2F3 = False And Form1.F6Continous = False Then                          'Gerbing 27.03.2014
            Call Form1.BildAnzeigen                                                             'Gerbing 27.03.2014
            frmGridAndThumb.Show                                                                'Gerbing 27.03.2014
            If frmGridAndThumb.DBGridNeu.SelBookmarks.Count = 1 Then                            'Gerbing 27.03.2014
                frmGridAndThumb.DBGridNeu.SelBookmarks.Remove 0                                 'Gerbing 27.03.2014
            End If                                                                              'Gerbing 27.03.2014
            frmGridAndThumb.DBGridNeu.SelBookmarks.Add frmGridAndThumb.rsDataGrid.Bookmark      'Gerbing 27.03.2014
        End If                                                                                  'Gerbing 27.03.2014
        'Ausgeben der Ausführungsdauer
        'glngEndMillisek = timeGetTime
        'Debug.Print "EndMillisec=" & glngendMillisek
        'Debug.Print "Millisekunden für RowColChange1" & "=" & (glngEndMillisek - glngStartMillisek)
'--------------------------------------------------------------------------------------------------
        'Falls es Thumbnails gibt, soll der zugehörige Thumbnail selektiert werden (blaue Umrandung)        'Gerbing 29.03.2015
        If blnComeFromBtnMitThumbnailsClick = True Then
            If gblnVollversion = False Then
                If gintDiffTage > 365 Then
                    If Sprache = 0 Then
                        MsgBox "Sie benutzen die Shareware-Version von GERBING Fotoalbum."      'Gerbing 25.06.2015
                    Else
                        MsgBox "You use the Shareware version of GERBING Fotoalbum."
                    End If
                    'Copy.Show 1                                                                'Gerbing 24.06.2015
                End If
            End If
            'Messen der Ausführungsdauer
            'glngStartMillisek = timeGetTime
            strTemp = Replace(gstrRowColChangeName, "+:\", gstrFotosMdbLocation & "\")
            blnThumbnailGefunden = False                                                        'Gerbing 10.11.2016
            'Dursuche alle Thumbnails ob es einen gleichnamigen Dateiname gibt
            For i = 0 To frmGridAndThumb.optThumb.Count - 1
                pos = InStr(1, frmGridAndThumb.Ulabel(i).Tag, strTemp, vbTextCompare)
                If pos <> 0 Then
                    blnThumbnailGefunden = True                                                 'Gerbing 10.11.2016
                    Exit For
                End If
            Next i
            If blnThumbnailGefunden = True Then                                                 'Gerbing 10.11.2016
                'Ausgeben der Ausführungsdauer
                'glngEndMillisek = timeGetTime
                'Debug.Print "EndMillisec=" & glngStartMillisek
                'Debug.Print "Millisekunden für RowColChange2" & "=" & (glngEndMillisek - glngStartMillisek)
    
                MerkeIndex = i
                glngGridline = MerkeIndex
                'das aktuelle soll in den sichtbaren Bereich gerückt werden und blau markiert werden
                'das wird durch den MerkeIndex bestimmt
                'Messen der Ausführungsdauer
                'glngStartMillisek = timeGetTime
                If blnCreateThumbnailsRunning = False Then
                    vsbSlide.Value = MerkeIndex
                    Call ChangePicFrameSize(MerkeIndex)
                End If
                'Ausgeben der Ausführungsdauer
                'glngEndMillisek = timeGetTime
                'Debug.Print "EndMillisec=" & glngendMillisek
                'Debug.Print "Millisekunden für RowColChange3" & "=" & (glngEndMillisek - glngStartMillisek)
            End If
        End If
'-----------------------------------------------------------------------------------------------
        gblnComeFromF2F3 = False                                                                'Gerbing 27.03.2014
    End If
    Exit Sub
RowColChangeError:
    msg = "Errornumber=" & Err.Number & vbNewLine
    msg = msg & Err.Description
    'MsgBox msg
    Resume Next
End Sub

Private Sub DbGridNeu_SelChange(Cancel As Integer)
    'Bei RowColChange stehen noch die Daten der Zeile zur Verfügung die den Fokus verloren hat
    'Bei SelChange stehen die Daten der Zeile zur Verfügung die den Fokus bekommen hat
    'SelChange kommt nur dran When the user selects a single row by clicking its record selector
End Sub

Private Sub Form_Terminate()
    Set ttToolTip = Nothing                                                    'Gerbing 06.06.2015
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim lIdx As Long
    
    SetWindowLong DBGridNeu.hWnd, GWL_WNDPROC, lpPrevWndProc                    'Gerbing 15.11.2019
    Me.Hide
    If gblnComefromVideo = True Then                                            'Gerbing 16.06.2012
        frmVideo.Show
    End If
    If Not gblnComeFromButtonF8 = True Then
        Cancel = True       'ich will kein Unload bei Klick aufs Schließkreuz, aber doch wenn F8 gedrückt war
                            'und unbedingt ohne cancel=true bei Beenden des Programms, sonst kommt Laufzeitfehler '91'
    Else
'        End                                                                     'Gerbing 06.06.2015
        If optThumb.Count > 2 Then
            For lIdx = 1 To optThumb.Count - 1
                On Error Resume Next
                Unload optThumb(lIdx)
                Unload Ulabel(lIdx)
                If Err.Number <> 0 Then Exit For
            Next lIdx
        End If
    End If
End Sub

Private Sub optThumb_GotFocus(Index As Integer)
'    If Index = 0 Then                       'Gerbing 04.05.2015 auskommentiert am 24.06.2015
'        Call optThumb_Click(Index)          'weil sonst ein Klick auf den ersten Thumbnail mit index=0 wirkungslos war
'    End If
End Sub

Private Sub optThumb_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single) 'Gerbing 29.10.2019
    If Button = vbRightButton Then
        Call Form1.Hilfebox                                                                                     'Gerbing 29.10.2019
    End If
End Sub

Private Sub pbBottom_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)                   'Gerbing 29.10.2019
    If Button = vbRightButton Then
        Call Form1.Hilfebox                                                                                     'Gerbing 29.10.2019
    End If
End Sub

Private Sub pbBottomLeer_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        Call Form1.Hilfebox                                                                                     'Gerbing 29.10.2019
    End If
End Sub

Private Sub pbTop_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)         ' Gerbing 29.03.2015 auskommentiert
End Sub

Function GetShellError(lErrorCode As Long) As String
    Const SE_ERR_FNF = 2&, SE_ERR_PNF = 3&
    Const SE_ERR_ACCESSDENIED = 5&, SE_ERR_OOM = 8&
    Const SE_ERR_DLLNOTFOUND = 32&, SE_ERR_SHARE = 26&
    Const SE_ERR_ASSOCINCOMPLETE = 27&, SE_ERR_DDETIMEOUT = 28&
    Const SE_ERR_DDEFAIL = 29&, SE_ERR_DDEBUSY = 30&
    Const SE_ERR_NOASSOC = 31&, ERROR_BAD_FORMAT = 11&

    Select Case lErrorCode
        Case SE_ERR_FNF
            GetShellError = "File not found"
        Case SE_ERR_PNF
            GetShellError = "Path not found"
        Case SE_ERR_ACCESSDENIED
            GetShellError = "Access denied"
        Case SE_ERR_OOM
            GetShellError = "Out of memory"
        Case SE_ERR_DLLNOTFOUND
            GetShellError = "DLL not found"
        Case SE_ERR_SHARE
            GetShellError = "A sharing violation occurred"
        Case SE_ERR_ASSOCINCOMPLETE
            GetShellError = "Incomplete or invalid file association"
        Case SE_ERR_DDETIMEOUT
            GetShellError = "DDE Time out"
        Case SE_ERR_DDEFAIL
            GetShellError = "DDE transaction failed"
        Case SE_ERR_DDEBUSY
            GetShellError = "DDE busy"
        Case SE_ERR_NOASSOC
            GetShellError = "No association for file extension"
        Case ERROR_BAD_FORMAT
            GetShellError = "Invalid EXE file or error in EXE image"
        Case Else
            GetShellError = "Unknown error"
    End Select
End Function

Public Sub SetSpaltenBreite()
    Dim ColWidth As Long
    Dim ColCaption As String
    Dim n As Long
    Dim Werte() As Long
    Dim AnzahlStandardfelder As Long
    
    AnzahlStandardfelder = 13
    ReDim Werte(1 + AnzahlStandardfelder + ND.ListNutzerdefinierteFelder.ListItems.Count)
    SQL = "SELECT SpaltenBreite.* FROM SpaltenBreite;"
    'SQL = "SELECT " & LoadResString(2525 + Sprache) & ".* FROM " & LoadResString(2525 + Sprache) & ";" 'Gerbing 08.11.2005
    On Error Resume Next
    rstsql.Close
    On Error GoTo 0
    With rstsql
        .Source = SQL
        .ActiveConnection = DBado                                           'Gerbing 23.11.2017
        .CursorType = adOpenForwardOnly
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    If Err.Number = 3078 Then
'        msg = "Seit Version 10.0.0.0 gibt es in der Datenbank fotos.mdb eine Tabelle SpaltenBreite," & NL
        msg = LoadResString(2014 + Sprache) & NL
'        msg = msg & "wo Änderungen des Nutzers an den Spaltenbreiten im Fenster Datenbank-Übersicht (siehe F5)" & NL
        msg = msg & LoadResString(2015 + Sprache) & NL
'        msg = msg & "eingetragen werden." & NL & NL
        msg = msg & LoadResString(2016 + Sprache) & NL & NL
'        msg = msg & "Diese Tabelle wurde nicht gefunden."
        msg = msg & LoadResString(2017 + Sprache)
        MsgBox msg                                     'Gerbing 08.11.2005
        rstsql.Close
        Exit Sub
    End If
    If rstsql.EOF Then
'        msg = "Seit Version 10.0.0.0 gibt es in der Datenbank fotos.mdb eine Tabelle SpaltenBreite," & NL
        msg = LoadResString(2014 + Sprache) & NL
'        msg = msg & "wo Änderungen des Nutzers an den Spaltenbreiten im Fenster Datenbank-Übersicht (siehe F5)" & NL
        msg = msg & LoadResString(2015 + Sprache) & NL
'        msg = msg & "eingetragen werden." & NL & NL
        msg = msg & LoadResString(2016 + Sprache) & NL & NL
'        msg = msg & "Diese Tabelle wurde nicht gefunden."
        msg = msg & LoadResString(2017 + Sprache)
        MsgBox msg                                     'Gerbing 08.11.2005
        rstsql.Close
        Exit Sub
    End If
    On Error GoTo 0

    n = 0
    ColCaption = DBGridNeu.Columns(0).Caption
'        If ColCaption = LoadResString(2524 + Sprache) Then                  'merker
'            n = 1
'        End If
    Do Until rstsql.EOF
        If n = DBGridNeu.Columns.Count Then Exit Do
        DBGridNeu.Columns(n).width = rstsql.Fields("Spaltenbreite")
        n = n + 1
        rstsql.MoveNext
    Loop
    rstsql.Close
End Sub

Public Sub SpeichernSpaltenBreite()
    Dim n As Long
    Dim ColWidth As Long
    
    SQL = "SELECT SpaltenBreite.* FROM SpaltenBreite;"
    'SQL = "SELECT " & LoadResString(2525 + Sprache) & ".* FROM " & LoadResString(2525 + Sprache) & ";" 'Gerbing 08.11.2005
    On Error Resume Next
    rstsql.Close
    On Error GoTo 0
    With rstsql
        .Source = SQL
        .ActiveConnection = DBado                                                       'Gerbing 23.11.2017
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With

    'Bei jedem Speichern der Spaltenbreiten wird der bisherige Inhalt der Tabelle Spaltenbreite zuerst
    'gelöscht, dann werden neue Einträge gemacht
    On Error GoTo 0
    
    If gblnSQLServerVersion = True Then
        'beim SQL Server muss es heißen 'Delete from table
        SQL = "DELETE From Spaltenbreite"
        'SQL = "DELETE FROM " & LoadResString(2525 + Sprache)
    Else
        SQL = "DELETE * From Spaltenbreite"
        'SQL = "DELETE * FROM " & LoadResString(2525 + Sprache)          '2525=Spaltenbreite
    End If
    DBado.Execute (SQL)                                                                 'Gerbing 23.11.2017
    For n = 0 To DBGridNeu.Columns.Count - 1                        'es wird ab Spalte Merker gespeichert
        rstsql.AddNew
        ColWidth = DBGridNeu.Columns(n).width
        If DBGridNeu.Columns(n).Visible = False Then ColWidth = 0
        rstsql.Fields("Spaltenbreite") = ColWidth
        rstsql.Update
    Next n
    rstsql.Close
End Sub

Private Sub CreateThumbPic(picSource As PictureBox, picThumb As PictureBox)
'This sub uses the halftone stretch mode, which produces the highest
'quality possible, when stretching the bitmap.

    Dim lRet            As Long
    Dim lLeft           As Long
    Dim lTop            As Long
    Dim lWidth          As Long
    Dim lHeight         As Long
    Dim lForeColor      As Long
    Dim hBrush          As Long
    Dim hDummyBrush     As Long
    Dim lOrigMode       As Long
    Dim fScale          As Single
    Dim uBrushOrigPt    As POINTAPI
    Dim NewFileName As String                                                           'Gerbing 10.11.2016
    Dim start As Long
    Dim pos As Long
    Dim DateinamenErweiterung As String
    
'    picThumb.Width = 110                 'Gehört zu optThumb.Width = 130 und Ulabel().width = 110 und Ulabel().Height = 31
'    picThumb.Height = 110               'Gehört zu optThumb.Height = 150

    'erproben wegen DPI 96                  'Gerbing 22.04.2015
    picThumb.width = 120
    If picThumb.width = 120 Then
        picThumb.height = 120
        optThumb(0).width = 130
        optThumb(0).height = 150
        Ulabel(0).width = 120
        Ulabel(0).height = 25
        'anpassen in ChangePicFrameSize - Ulabel(lIdx).Move x + 5, y + 10
    End If


'    picThumb.Width = 100                 'Gehört zu optThumb.Width = 120
'    picThumb.Height = 100                'Gehört zu optThumb.Height = 140
'    picThumb.Width = 90                 'Gehört zu optThumb.Width = 100
'    picThumb.Height = 90                'Gehört zu optThumb.Height = 120
'    picThumb.Width = 64                 'Gehört zu optThumb.Width = 80
'    picThumb.Height = 64                'Gehört zu optThumb.Height = 110
'    picThumb.Width = 180               'Gehört zu optThumb.Width = 200
'    picThumb.Height = 180              'Gehört zu optThumb.Height = 240

    picThumb.BackColor = vbButtonFace
    picThumb.AutoRedraw = True
    picThumb.Cls
    
    If picSource.width <= picThumb.width - 2 And picSource.height <= picThumb.height - 2 Then
        fScale = 1
    Else
        fScale = IIf(picSource.width > picSource.height, (picThumb.width - 2) / picSource.width, (picThumb.height - 2) / picSource.height)
    End If
    lWidth = picSource.width * fScale
    lHeight = picSource.height * fScale
    lLeft = Int((picThumb.width - lWidth) / 2)
    lTop = Int((picThumb.height - lHeight) / 2)
    
    'Store the original ForeColor
    lForeColor = picThumb.ForeColor
    
    'Set picEdit's stretch mode to halftone (this may cause misalignment of the brush)
    lOrigMode = SetStretchBltMode(picThumb.hDC, STRETCH_HALFTONE)
    
    'Realign the brush...
    'Get picEdit's brush by selecting a dummy brush into the DC
    hDummyBrush = CreateSolidBrush(lForeColor)
    hBrush = SelectObject(picThumb.hDC, hDummyBrush)
    'Reset the brush (This will force windows to realign it when it's put back)
    lRet = UnrealizeObject(hBrush)
    'Set picEdit's brush alignment coordinates to the left-top of the bitmap
    lRet = SetBrushOrgEx(picThumb.hDC, lLeft, lTop, uBrushOrigPt)
    'Now put the original brush back into the DC at the new alignment
    hDummyBrush = SelectObject(picThumb.hDC, hBrush)
    
    'Stretch the bitmap
    lRet = StretchBlt(picThumb.hDC, lLeft, lTop, lWidth, lHeight, _
            picSource.hDC, 0, 0, picSource.width, picSource.height, SRCCOPY)
    
    'Set the stretch mode back to it's original mode
    lRet = SetStretchBltMode(picThumb.hDC, lOrigMode)
    
    'Reset the original alignment of the brush...
    'Get picEdit's brush by selecting the dummy brush back into the DC
    hBrush = SelectObject(picThumb.hDC, hDummyBrush)
    'Reset the brush (This will force windows to realign it when it's put back)
    lRet = UnrealizeObject(hBrush)
    'Set the brush alignment back to the original coordinates
    lRet = SetBrushOrgEx(picThumb.hDC, uBrushOrigPt.x, uBrushOrigPt.y, uBrushOrigPt)
    'Now put the original brush back into picEdit's DC at the original alignment
    hDummyBrush = SelectObject(picThumb.hDC, hBrush)
    'Get rid of the dummy brush
    lRet = DeleteObject(hDummyBrush)
    
    'Restore the original ForeColor
    picThumb.ForeColor = lForeColor
    'picThumb.Line (lLeft - 1, lTop - 1)-Step(lWidth + 1, lHeight + 1), &H0&, B
End Sub

Private Sub CreateThumbs()
    Dim iMaxLen As Integer
    Dim lIdx    As Long
    Dim lPicCnt As Long
    Dim lFilCnt As Long
    Dim sText   As String
    Dim TeilFile As String
    Dim ThumbFileNameFoto As String
    Dim start As Long
    Dim pos As Long
    Dim DateinamenErweiterung As String
    Dim strPath As String
    Dim strFile As String
    Dim ExtProp As Variant
    Dim GdipLoader As GdipLoader
    
    'Dim folder As Variant                                                'Gerbing 30.11.2016
    
    Dim ShellFolderItem As shell32.ShellFolderItem
    'Dim ShellFolderItem As Variant                                       'Gerbing 30.11.2016
    
    Dim PropLong As Variant
    Dim PropStream As IUnknown
    Dim outname As String
    Dim blnFN As Boolean
    Dim CreateThumbsZähler As Long
    Dim AnzahlSpalten As Long


    picFrame.Visible = False
    Screen.MousePointer = vbHourglass
    Do While optThumb.Count > 2
        On Error Resume Next
        Unload optThumb(optThumb.Count - 1)
        Unload Ulabel(optThumb.Count - 1)
        If Err.Number <> 0 Then Exit Do
    Loop
    On Error GoTo 0
    TimerfrmGridAndThumb.Enabled = True
    DoEvents
    On Error Resume Next
    If Koll.Count > 0 Then
        ProgBar.Max = Koll.Count
        KollIndex = 1
        blnAbbrechen = False
        For lIdx = 0 To Koll.Count - 1
            'Schleife durch alle gefundenen Dateienamen
            StatBarU.Panels(0).Text = Koll.Item(KollIndex)
            Set picLoad.Picture = LoadPicture()
            picLoad.Cls
            Err.Clear
            '-------------------------------------
            Do
                'Das gehört nach fotos.exe                                                      'Gerbing 10.11.2016 Start
                'Wenn es den Thumbnail im Ordner ...\GerbingThumbs\... gibt, nehme ich diesen
                ThumbFileNameFoto = Koll.Item(KollIndex)                                        'Gerbing 10.11.2016 Start Gerbing 07.12.2016
                start = 1
                Do
                    pos = InStr(start, ThumbFileNameFoto, "\")
                    If pos = 0 Then Exit Do
                    start = pos + 1
                Loop
                ThumbFileNameFoto = Left(ThumbFileNameFoto, start - 1) & "GerbingThumbs\" & Right(ThumbFileNameFoto, Len(ThumbFileNameFoto) - start + 1) & ".jpg"
                blnFN = False
                If file_path_exist(ThumbFileNameFoto) = True Then
                    Set picLoad.Picture = CreateThumbnailFromFile(ThumbFileNameFoto, 100)               'ich nehme das aus GerbingThumbs
                    Exit Do
                Else
                    'wenn es den Thumbnail in ...\GerbingThumbs\ nicht gibt wird dort einer erzeugt
                    '-------------------------------------
                    'file_split splits a complete file name into directory, file name and extension:
                    Call file_split(Koll.Item(KollIndex), strPath, strFile, DateinamenErweiterung)
                    DateinamenErweiterung = UCase(DateinamenErweiterung)
                    Select Case DateinamenErweiterung
                        Case "AVI", "MPG", "MOV", "WMV", "ASF", "MP4", "ASX", "MKV", "FLV"      'Gerbing 10.12.2017
                        'Case "AVI", "MPG", "MOV", "WMV", "ASF"                          'bei Videos
                            Set folder = ShellObject.NameSpace(strPath)
                            'Set folder = CreateObject("Shell.Application").NameSpace(strPath)
                            'Set folder = ShellObject.NameSpace(CVar(strPath)) 'argument must be a variant   'gerbing 30.11.2016
                            If folder Is Nothing Then
                                If gblnDieseNachrichtNichtMehrZeigen = False Then
                                    'MsgBox "Folder Is Nothing"
                                    frmFolderIsNothing.Show 1
                                End If
                                Exit Do                                                                     'Gerbing 30.11.2016
                            Else
                                Set GdipLoader = New GdipLoader
                            End If
                            Set ShellFolderItem = folder.ParseName(strFile & "." & DateinamenErweiterung)
                            PropLong = ShellFolderItem.ExtendedProperty(SCID_PerceivedType)
                            If Not IsEmpty(PropLong) Then
                                If PropLong = PERCEIVED_TYPE_VIDEO Then
                                    On Error Resume Next                                                    'Gerbing 24.11.2016
                                    Set PropStream = ShellFolderItem.ExtendedProperty(SCID_PropStream)
                                    If Err.Number = 0 Then
                                        If Not PropStream Is Nothing Then
                                            'outname = strPath & "\GerbingThumbs\" & ShellFolderItem.Name & ".jpg"
                                            outname = strPath & "GerbingThumbs\" & ShellFolderItem.Name & ".jpg" 'Gerbing 21.01.2018
                                            On Error Resume Next
                                            If file_path_exist(strPath & "GerbingThumbs\") = False Then
                                                MkDir strPath & "GerbingThumbs\"                            'Gerbing 21.01.2018
                                                'MkDir ersetzen durch CreateDirectoryW                      'Gerbing 21.01.2018
                                            End If
                                            On Error GoTo 0
                                            'erzeuge video thumbnail
                                            GdipTool.PropStream2PicFileScaled PropStream, outname, PFF_JPEG 'Gerbing 24.11.2016
                                            'GdipTool.PropStream2PicFileScaled PropStream, outname, PFF_JPEG, , 130, 90 'Gerbing 05.06.2017
                                            'anschließend geht es im Loop wieder nach oben und benutzt den soeben erzeugte thumbnail
                                        End If
                                    Else
                                        'für dieses video konnte kein thumbnail erzeugt werden drum kann ich den Loop verlassen
                                        Exit Do
                                    End If
                                End If
                                On Error GoTo 0                                                         'Gerbing 24.11.2016
                            End If
                        Case Else
                            '--------------------------------------------------
                            'bei JPG und anderen Bildern und bei MP4 und ASX
                            'vielleicht geht es mit MP4 und ASX ja auch wenn andere Codecs installiert sind
                            'bei http://www.vbforums.com/showthread.php?761717-VB6-Shell-Video-Thumbnail-Images nachlesen
                            outname = strPath & "GerbingThumbs\" & strFile & "." & DateinamenErweiterung & ".jpg" 'Gerbing 07.12.2016
                            DateinamenErweiterung = UCase(DateinamenErweiterung)
                            Set picLoad.Picture = CreateThumbnailFromFile(Koll.Item(KollIndex), 25)
                            Call CreateThumbPic(picLoad, picThumb)
                            If file_path_exist(strPath & "GerbingThumbs\") = False Then
                                MkDir strPath & "GerbingThumbs\"                            'Gerbing 21.01.2018
                                'MkDir ersetzen durch CreateDirectoryW                      'Gerbing 21.01.2018
                            End If
                             'erzeuge JPG thumbnail
                            PicSave.SavePicture picThumb.Image, outname, fmtJPEG, 70            'Gerbing 10.11.2016 End
                            '--------------------------------------------------
                    End Select
                    DoEvents
                End If
            Loop
            Call CreateThumbPic(picLoad, picThumb)
            If CreateThumbsZähler > 0 Then                                          'beim ersten mal ist CreateThumbsZähler=0
                Load optThumb(CreateThumbsZähler)
                Set optThumb(CreateThumbsZähler).Container = picFrame
                Load Ulabel(CreateThumbsZähler)
            End If
            TeilFile = file_sep(Koll.Item(KollIndex), True)
            optThumb(CreateThumbsZähler).Tag = Koll.Item(KollIndex)

            Set optThumb(CreateThumbsZähler).Picture = picThumb.Image
            sText = TeilFile
            iMaxLen = optThumb(CreateThumbsZähler).width * 2                        'Gerbing 29.03.2015
            If iMaxLen < optThumb(CreateThumbsZähler).width - 15 Then
                sText = sText & "..."
            End If
            Ulabel(CreateThumbsZähler).Text = sText
            Ulabel(CreateThumbsZähler).Tag = Koll.Item(KollIndex)                   'Tag enthält den ganzen Dateiname
            optThumb(CreateThumbsZähler).Visible = False
            Ulabel(CreateThumbsZähler).Visible = False
            'EndMillisek = timeGetTime
            'Ausgeben der Ausführungsdauer
'            Debug.Print "EndMillisec=" & EndMillisek
'            Debug.Print "Millisekunden für Datei" & Koll.Item(KollIndex) & "=" & (EndMillisek - StartMillisek)
            KollIndex = KollIndex + 1
            CreateThumbsZähler = CreateThumbsZähler + 1
            DoEvents
        Next lIdx
        'Free the unneeded resources
        Set picLoad.Picture = LoadPicture()
        Set picThumb.Picture = LoadPicture()
        Screen.MousePointer = vbDefault                                             'Gerbing 24.06.2015
        If blnAbbrechen = False Then
            optThumb(0).Value = True
        End If
        'mlCurThumb = 0                                                             'Gerbing 29.03.2015 auskommentiert
        glngGridline = 1
        Call Form_Resize
        picFrame.Visible = True
        TimerfrmGridAndThumb.Enabled = False
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub optThumb_Click(Index As Integer)
    Dim i As Long
    Dim txtSuchen As String
    Dim strFind As String
    Dim lngTemp As Long                                                                 'Gerbing 10.11.2016
    Dim DateinamenErweiterung As String
    Dim intLänge As Long
    
    If blnComeFromSlideChange = True Then Exit Sub                                      'Gerbing 27.11.2016
    If Index = 0 Then
        Exit Sub                                                                        'Gerbing 10.11.2016
    End If
    If Index >= Koll.Count Then                                                         'Gerbing 10.11.2016
        Index = Koll.Count - 1
    End If
    If gblnVollversion = False Then
        If gintDiffTage > 365 Then
            Copy.Show 1                                                                 'Gerbing 24.06.2015
        End If
    End If
    'If optThumb.Count > 2 Then
        On Error Resume Next
        For i = 0 To optThumb.Count - 1
            optThumb(i).BackColor = vbButtonFace
            optThumb(i).tooltipText = ""                                                'Gerbing 09.06.2015
        Next i
        On Error GoTo 0
    'End If
    optThumb(Index).BackColor = vbBlue
    If Not isIDE Then                                           'sonst ist debuggen fast unmöglich
        Set ttToolTip = New clsToolTip
        With ttToolTip
            .Create Me.hWnd
            .Activate
            '.AddTool cmd.hWnd, "Hello World! " & ChrW(&HFA23), , , False
            '.AddTool Ulabel(Index).hWnd, "hier tooltip auf Picture", , , False
            .AddTool optThumb(Index).hWnd, Ulabel(Index).Tag, , , False
        End With
    Else
        optThumb(Index).tooltipText = Ulabel(Index).Tag                                     'Gerbing 08.06.2015 09.06.2015
    End If
    gblnComeFromThumbs = True
    gblnWasOptThumbClick = True                                                             'Gerbing 04.05.2015
    gstrFRODN = Ulabel(Index).Tag
    gblnfrmGridAndThumbDblClick = True                                                      'Gerbing 30.11.2016
    
    Call Form1.BildAnzeigen
    
    gblnComeFromThumbs = False
    'Im Grid soll derselbe Satz eingestellt werden, wie der dem geklickten Thumbnail entsprechende
    ' Der Dateiname muss verwandelt werden in die Form wie er in der Datenbank steht
    gstrRowColChangeName = Ulabel(Index).Tag
    txtSuchen = Ulabel(Index).Tag
    txtSuchen = Replace(txtSuchen, gstrFotosMdbLocation & "\", "+:\")
    'wenn im Dateinamen ein "'" vorkommt, ersetzen durch 2 Hochkommas                       'Gerbing 23.01.2018
    txtSuchen = Replace(txtSuchen, "'", "''")                                               'Gerbing 23.01.2018
    'strFind = "Dateiname like '*" & txtSuchen & "*'"
    strFind = LoadResString(1028 + Sprache) & " like '*" & txtSuchen & "*'"
    On Error GoTo 0
    'hier wird zweimal frmGridAndThumb_RowColChange aufgerufen, bei MoveFirst und bei Find und in der Folge die Function Form1.Bildanzeigen
    'das will ich verhindern weil es flackert
    gblnComeFromThumbs = True
    frmGridAndThumb.Adodc1.Recordset.MoveFirst                   'unbedingt vor ...Find musst Du .MoveFirst machen
    frmGridAndThumb.Adodc1.Recordset.Find strFind                'sonst kommt error 3021 wenn der Satz weiter vorn steht als der aktuelle
    gblnComeFromThumbs = False
    If frmGridAndThumb.DBGridNeu.SelBookmarks.Count = 1 Then
        frmGridAndThumb.DBGridNeu.SelBookmarks.Remove 0
    End If
    frmGridAndThumb.DBGridNeu.SelBookmarks.Add frmGridAndThumb.rsDataGrid.Bookmark
    If Index > 0 Then
        On Error Resume Next   'sonst kommt Laufzeitfehler '380'                            'Gerbing 16.11.2016
        lngTemp = optThumb(Index)
        vsbSlide.Value = lngTemp                                                            'Gerbing 16.11.2016
        On Error GoTo 0                                                                     'Gerbing 16.11.2016
    End If
    '---------------------------------------------------------------------------------------'Gerbing 30.11.2016
    'nur bei Wechsel von Video auf Foto muss Form1 sichtbar gemacht werden
    DateinamenErweiterung = Ulabel(Index).Tag
    If Mid(DateinamenErweiterung, Len(DateinamenErweiterung) - 3, 1) = "." Then             'Gerbing 25.06.2006
        intLänge = 3
    End If
    If Mid(DateinamenErweiterung, Len(DateinamenErweiterung) - 4, 1) = "." Then
        intLänge = 4
    End If
    If Mid(DateinamenErweiterung, Len(DateinamenErweiterung) - 5, 1) = "." Then
        intLänge = 5
    End If
    DateinamenErweiterung = Right(gstrRowColChangeName, intLänge)
    DateinamenErweiterung = UCase(DateinamenErweiterung)            'Gerbing 26.11.2006
    Select Case DateinamenErweiterung
        Case "AVI", "MPG", "PEG", "MOV", "MPE", "ASF", "ASX", "WMV", "MP4", "MKV", "FLV"      'Gerbing 10.12.2017
            'Me.Show                                                                        'Gerbing 27.11.2016
        Case Else
            Form1.Show
    End Select
End Sub

Private Sub optThumb_DblClick(Index As Integer)
    Dim i As Long
    Dim txtSuchen As String
    Dim strFind As String
    
    Me.Hide
    gblnComeFromThumbs = True
    gblnWasOptThumbClick = True                                                         'Gerbing 04.05.2015
    gstrFRODN = Ulabel(Index).Tag
    Call Form1.BildAnzeigen
    gblnComeFromThumbs = False
    
    'Im Grid soll derselbe Satz eingestellt werden, wie der dem geklickten Bild entsprechende
    ' Der Dateiname muss verwandelt werden in die Form wie er in der Datenbank steht
    txtSuchen = Replace(gstrFRODN, gstrFotosMdbLocation & "\", "+:\")
    txtSuchen = Replace(txtSuchen, "'", "''")                                           'Gerbing 23.01.2018
    'strFind = "Dateiname like '*" & txtSuchen & "*'"
    strFind = LoadResString(1028 + Sprache) & " like '*" & txtSuchen & "*'"
    'wenn im Dateinamen ein "'" vorkommt, ersetzen durch 2 Hochkommas                   'Gerbing 23.01.2018
    On Error GoTo 0
    'hier wird zweimal frmGridAndThumb_RowColChange aufgerufen, bei MoveFirst und bei Find und in der Folge die Function Form1.Bildanzeigen
    'das will ich verhindern weil es flackert
    gblnComeFromThumbs = True
    frmGridAndThumb.Adodc1.Recordset.MoveFirst                   'unbedingt vor ...Find musst Du .MoveFirst machen
    frmGridAndThumb.Adodc1.Recordset.Find strFind                'sonst kommt error 3021 wenn der Satz weiter vorn steht als der aktuelle
    gblnComeFromThumbs = False
    If frmGridAndThumb.DBGridNeu.SelBookmarks.Count = 1 Then
        frmGridAndThumb.DBGridNeu.SelBookmarks.Remove 0
    End If
    frmGridAndThumb.DBGridNeu.SelBookmarks.Add frmGridAndThumb.rsDataGrid.Bookmark
End Sub

Private Sub picFrame_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        Call Form1.Hilfebox                                                                                     'Gerbing 29.10.2019
    End If
End Sub

Private Sub TimerfrmGridAndThumb_Timer()
  On Error Resume Next
  ProgBar.Value = KollIndex
  On Error GoTo 0
End Sub

Private Sub KollFüllen()
    Dim MyRecordset As ADODB.Recordset
    Dim Filename As String
    
    Screen.MousePointer = vbHourglass
    gblnSubdirectories = False
    'In GERBING Fotoalbum wird anstelle von Function Rekursive der von der Abfrage gelieferte recordset benutzt
    'Call Rekursive(FolderPath, "*.*")
    
    ' Recordset erstellen und öffnen adOpenStatic
    Set MyRecordset = New ADODB.Recordset
    With MyRecordset
        If gblnWasHeadClick = False Then
            .Source = frmGridAndThumb.Adodc1.RecordSource
        Else
            .Source = frmGridAndThumb.SQLneuHeadClick
        End If
        .ActiveConnection = DBado                                                       'Gerbing 23.11.2017
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    While Not MyRecordset.EOF
        Filename = Replace(MyRecordset.Fields(LoadResString(1028 + Sprache)), "+:\", gstrFotosMdbLocation & "\")  '1028=Dateiname  'Gerbing 07.11.2011
        Koll.Add Filename                                                           'Gerbing 24.11.2016

        MyRecordset.MoveNext
    Wend
    MyRecordset.Close
    Screen.MousePointer = vbNormal                                                  'Gerbing 07.01.2009
End Sub

Private Sub StatBarU_ResizedControlWindow()
  Dim X1 As Long
  Dim X2 As Long
  Dim Y1 As Long
  Dim Y2 As Long

  StatBarU.Panels(1).GetRectangle X1, Y1, X2, Y2
  MoveWindow ProgBar.hWnd, X1, Y1 + 1, X2 - X1 - 4, Y2 - Y1 - 2, 1
End Sub

Private Sub ChangePicFrameSize(Gridline As Long)
    'Gridline ist das aktuelle Bild in DbGridNeu. Dieses Bild will ich in PicFrame an eine mittlere Position stellen.
    Dim i As Long
    Dim x       As Long
    Dim y       As Long
    Dim lIdx    As Long
    Dim lCols   As Long
    Dim lRows   As Long
    Dim lAnzahlThumbsSichtbar As Long
    Dim startThumb As Long
        
    If Gridline >= Koll.Count Then                                                                              'Gerbing 10.11.2016
        Gridline = Koll.Count - 1
    End If
    If Me.WindowState <> vbMinimized Then
        If Me.width < 346 * Screen.TwipsPerPixelX Then
            Me.width = 346 * Screen.TwipsPerPixelX
        ElseIf Me.height < 378 * Screen.TwipsPerPixelY Then
            Me.height = 378 * Screen.TwipsPerPixelY
        Else
            If UBound(TV) <> 0 Then                                                'die zuvor gezeigten unsichtbar machen
                For i = 0 To UBound(TV)
                    optThumb(TV(i)).Visible = False
                    Ulabel(TV(i)).Visible = False
                Next i
            End If
            If blnComeFromBtnMitThumbnailsClick = True And Koll.Count > 0 Then
                On Error Resume Next
                picFrame.Move 0, 0, pbBottom.ScaleWidth - vsbSlide.width, pbBottom.ScaleHeight - StatBarU.height - 30
                vsbSlide.Move pbBottom.ScaleWidth - vsbSlide.width - 5, 0, vsbSlide.width, picFrame.ScaleHeight  'Gerbing 10.11.2016
                On Error GoTo 0
                StatBarU.Move 0, pbBottom.ScaleHeight - StatBarU.height, pbBottom.ScaleWidth, StatBarU.height
                lCols = Int((picFrame.ScaleWidth) / optThumb(0).width)              'Anzahl Spalten
                lRows = Int((picFrame.ScaleHeight) / optThumb(0).height)            'Anzahl Zeilen
                lAnzahlThumbsSichtbar = lCols * lRows
                If optThumb.Count < lAnzahlThumbsSichtbar Then
                    lAnzahlThumbsSichtbar = optThumb.Count
                End If
                'Bisher liegt sowohl die Liste optThumb(index) wie auch die Liste Ulabel(index) ein Element über dem anderen
                If optThumb.Count >= 1 Then
                    
                    'Messen der Ausführungsdauer
                    'glngStartMillisek = timeGetTime
                    startThumb = Gridline - (lAnzahlThumbsSichtbar \ 2)
                    If startThumb + lAnzahlThumbsSichtbar - 1 >= optThumb.Count Then
                        startThumb = optThumb.Count - lAnzahlThumbsSichtbar
                    End If
                    If startThumb < 0 Then
                        startThumb = 0
                    End If
                    
                    ReDim TV(lAnzahlThumbsSichtbar)                                 'ich merke mir den jeweiligen index von optThumb und Ulabel
                    For i = 0 To lAnzahlThumbsSichtbar - 1                          'im Array TV()
                        TV(i) = startThumb + i
                    Next i
                    
                    For lIdx = startThumb To startThumb + lAnzahlThumbsSichtbar - 1
                        'jetzt wird jedes Element der Liste optThumb(index) wie auch der Liste Ulabel(index) an seinen eigenen Platz
                        'innerhalb von picFrame verschoben
                        'danach werden die betroffenen optThumb(index) und Ulabel(index) wieder sichtbar gemacht
                        x = ((lIdx - startThumb) Mod lCols) * optThumb(0).width
                        y = Int((lIdx - startThumb) / lCols) * optThumb(0).height
                        optThumb(lIdx).Move x, y
                        optThumb(lIdx).Visible = True
                        'Ulabel(lIdx).Move x + 10, y + 10
                        Ulabel(lIdx).Move x + 5, y + 10                                     'Gerbing 22.04.2015
                        Ulabel(lIdx).ZOrder
                        Ulabel(lIdx).Visible = True
                        'picFrame.Width = lCols * optThumb(0).Width
                        picFrame.width = lCols * optThumb(0).width + vsbSlide.width         'Gerbing 10.11.2016
                    Next lIdx
                    For lIdx = startThumb To startThumb + lAnzahlThumbsSichtbar - 1
                        optThumb(lIdx).BackColor = vbButtonFace                             'etwa schon markierte ausschalten
                    Next lIdx
                    frmGridAndThumb.optThumb(Gridline).BackColor = vbBlue                   'das gefundene wird blau
                    'Ausgeben der Ausführungsdauer
                    'glngEndMillisek = timeGetTime
                    'Debug.Print "EndMillisec=" & glngendMillisek
                    'Debug.Print "Millisekunden für ChangePicFrameSize3" & "=" & (glngEndMillisek - glngStartMillisek)
                End If
            End If
        End If
    End If
End Sub

Private Sub vsbSlide_Change()                                                       'Gerbing 10.11.2016
    Call ChangePicFrameSize(vsbSlide.Value)
    blnComeFromSlideChange = True                                                   'Gerbing 27.11.2016
    optThumb_Click (vsbSlide.Value)
    blnComeFromSlideChange = False                                                  'Gerbing 27.11.2016
End Sub

Private Sub vsbSlide_Scroll()                                                       'Gerbing 10.11.2016
    vsbSlide_Change
End Sub

