VERSION 5.00
Begin VB.Form frmZwischenablageOderOrdner 
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   2184
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   6852
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2184
   ScaleWidth      =   6852
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.PictureBox Picture1 
      Height          =   732
      Left            =   120
      ScaleHeight     =   684
      ScaleWidth      =   804
      TabIndex        =   4
      Top             =   1440
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   492
      Left            =   2520
      TabIndex        =   3
      Top             =   1560
      Width           =   1692
   End
   Begin VB.Frame Frame1 
      Height          =   1332
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6612
      Begin VB.OptionButton optOrdner 
         Caption         =   "Das Bild soll in einen Ordner Kopiert werden"
         Height          =   372
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   6252
      End
      Begin VB.OptionButton optZwischenablage 
         Caption         =   "Das Bild soll in die Zwischenablage"
         Height          =   372
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   6252
      End
   End
End
Attribute VB_Name = "frmZwischenablageOderOrdner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    '------------------------------------------------------------Gerbing 12.08.2017 für Strg+C
    ' Required data structures
    Private Type POINTAPI
    x As Long
    y As Long
    End Type
    
    ' Clipboard Manager Functions
    Private Declare Function EmptyClipboard Lib "user32" () As Long
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
    
    ' Other required Win32 APIs
    Private Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
    Private Declare Function DragQueryPoint Lib "shell32.dll" (ByVal hDrop As Long, lpPoint As POINTAPI) As Long
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    
    ' Predefined Clipboard Formats
    Private Const CF_TEXT = 1
    Private Const CF_BITMAP = 2
    Private Const CF_METAFILEPICT = 3
    Private Const CF_SYLK = 4
    Private Const CF_DIF = 5
    Private Const CF_TIFF = 6
    Private Const CF_OEMTEXT = 7
    Private Const CF_DIB = 8
    Private Const CF_PALETTE = 9
    Private Const CF_PENDATA = 10
    Private Const CF_RIFF = 11
    Private Const CF_WAVE = 12
    Private Const CF_UNICODETEXT = 13
    Private Const CF_ENHMETAFILE = 14
    Private Const CF_HDROP = 15
    Private Const CF_LOCALE = 16
    Private Const CF_MAX = 17
    
    ' New shell-oriented clipboard formats
    Private Const CFSTR_SHELLIDLIST As String = "Shell IDList Array"
    Private Const CFSTR_SHELLIDLISTOFFSET As String = "Shell Object Offsets"
    Private Const CFSTR_NETRESOURCES As String = "Net Resource"
    Private Const CFSTR_FILEDESCRIPTOR As String = "FileGroupDescriptor"
    Private Const CFSTR_FILECONTENTS As String = "FileContents"
    Private Const CFSTR_FILENAME As String = "FileName"
    Private Const CFSTR_PRINTERGROUP As String = "PrinterFriendlyName"
    Private Const CFSTR_FILENAMEMAP As String = "FileNameMap"
    
    ' Global Memory Flags
    Private Const GMEM_FIXED = &H0
    Private Const GMEM_MOVEABLE = &H2
    Private Const GMEM_NOCOMPACT = &H10
    Private Const GMEM_NODISCARD = &H20
    Private Const GMEM_ZEROINIT = &H40
    Private Const GMEM_MODIFY = &H80
    Private Const GMEM_DISCARDABLE = &H100
    Private Const GMEM_NOT_BANKED = &H1000
    Private Const GMEM_SHARE = &H2000
    Private Const GMEM_DDESHARE = &H2000
    Private Const GMEM_NOTIFY = &H4000
    Private Const GMEM_LOWER = GMEM_NOT_BANKED
    Private Const GMEM_VALID_FLAGS = &H7F72
    Private Const GMEM_INVALID_HANDLE = &H8000
    Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
    Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
    
    Private Type DROPFILES
    pFiles As Long
    pt As POINTAPI
    fNC As Long
    fWide As Long
    End Type


Private Sub Command1_Click()

End Sub

Private Sub btnOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call AnpassenNutzerWunsch(Me)                               'Gerbing 11.03.2017
    optZwischenablage.Caption = LoadResString(2350 + Sprache)   'Das Bild soll in die Zwischenablage
    optOrdner.Caption = LoadResString(2351 + Sprache) 'Das Bild soll in einen Ordner kopiert werden
End Sub

Private Sub optOrdner_Click()
    ClipboardCopyFiles (gblStrAktuellGezeigtesBild)
End Sub

Private Sub optZwischenablage_Click()
    'Picture1.Picture = LoadPicture(gblStrAktuellGezeigtesBild)
    
    'Picture1.Picture = LoadPicture(gblStrAktuellGezeigtesBild)                  'versteht nicht unicode
    Call LoadPictureWThumb(gblStrAktuellGezeigtesBild)                           'versteht unicode Gerbing 11.08.2017

    
    Clipboard.Clear
    Clipboard.SetData Picture1.Picture, vbCFDIB
End Sub

'Public Function ClipboardCopyFiles(Files() As String) As Boolean
Public Function ClipboardCopyFiles(file As String) As Boolean
    Dim Data As String
    Dim df As DROPFILES
    Dim hGlobal As Long
    Dim lpGlobal As Long
    Dim i As Long
    
    ' Open and clear existing crud off clipboard.
    If OpenClipboard(0&) Then
    Call EmptyClipboard
    '' Build double-null terminated list of files.
    'For i = LBound(Files) To UBound(Files)
    'data = data & Files(i) & vbNullChar
    'Next
    'data = data & vbNullChar
    ' Build double-null terminated list of files.
    Data = Data & file & vbNullChar
    Data = Data & vbNullChar
    ' Allocate and get pointer to global memory,
    ' then copy file list to it.
    hGlobal = GlobalAlloc(GHND, Len(df) + Len(Data))
    If hGlobal Then
    lpGlobal = GlobalLock(hGlobal)
    ' Build DROPFILES structure in global memory.
    df.pFiles = Len(df)
    Call CopyMem(ByVal lpGlobal, df, Len(df))
    Call CopyMem(ByVal (lpGlobal + Len(df)), ByVal Data, Len(Data))
    Call GlobalUnlock(hGlobal)
    ' Copy data to clipboard, and return success.
    If SetClipboardData(CF_HDROP, hGlobal) Then
    ClipboardCopyFiles = True
    End If
    End If
    ' Clean up
    Call CloseClipboard
    End If
End Function

