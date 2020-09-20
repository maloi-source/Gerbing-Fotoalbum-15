VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{B6CC61F6-3F1A-4B00-9918-13F66F185263}#1.0#0"; "LblCtlsU.ocx"
Object = "{A10D6B26-9A8F-4A87-A2D1-1D8C9EED0967}#1.3#0"; "StatBarU.ocx"
Begin VB.Form frmThumbs 
   Caption         =   "Thumb View"
   ClientHeight    =   5076
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   12984
   Icon            =   "Thumbs.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   423
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1082
   StartUpPosition =   2  'Bildschirmmitte
   WindowState     =   2  'Maximiert
   Begin VB.Timer TimerFrmThumbs 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   7560
      Top             =   720
   End
   Begin MSComctlLib.ProgressBar Progbar 
      Height          =   252
      Left            =   5760
      TabIndex        =   8
      Top             =   3840
      Width           =   3492
      _ExtentX        =   6160
      _ExtentY        =   445
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdAbbrechen 
      Caption         =   "Abbrechen"
      Height          =   372
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1932
   End
   Begin VB.PictureBox picThumb 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   5760
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   2
      Top             =   2400
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picLoad 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   1560
      Left            =   5760
      ScaleHeight     =   130
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox picFrame 
      BorderStyle     =   0  'Kein
      Height          =   3840
      Left            =   120
      ScaleHeight     =   320
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   247
      TabIndex        =   0
      Top             =   720
      Width           =   2964
      Begin VB.VScrollBar vsbSlide 
         Height          =   3060
         Left            =   2760
         TabIndex        =   5
         Top             =   0
         Width           =   210
      End
      Begin VB.PictureBox picSlide 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   3360
         Left            =   0
         ScaleHeight     =   280
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   230
         TabIndex        =   3
         Top             =   120
         Width           =   2760
         Begin LblCtlsLibUCtl.WindowedLabel Ulabel 
            Height          =   372
            Index           =   0
            Left            =   240
            TabIndex        =   6
            Top             =   1320
            Width           =   2172
            _cx             =   3831
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
            Text            =   "Thumbs.frx":038A
         End
         Begin VB.OptionButton optThumb 
            Height          =   2880
            Index           =   0
            Left            =   120
            Style           =   1  'Grafisch
            TabIndex        =   4
            Top             =   240
            Width           =   2400
         End
      End
   End
   Begin StatBarLibUCtl.StatusBar StatBarU 
      Align           =   2  'Unten ausrichten
      Height          =   348
      Left            =   0
      Top             =   4728
      Width           =   12984
      Version         =   258
      _cx             =   22902
      _cy             =   614
      Appearance      =   0
      BackColor       =   -1
      BorderStyle     =   0
      CustomCapsLockText=   "Thumbs.frx":03B6
      CustomInsertKeyText=   "Thumbs.frx":03DE
      CustomKanaLockText=   "Thumbs.frx":0404
      CustomNumLockText=   "Thumbs.frx":042C
      CustomScrollLockText=   "Thumbs.frx":0452
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
            Text            =   "Thumbs.frx":047A
            Object.ToolTipText     =   "Thumbs.frx":04A4
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
            Text            =   "Thumbs.frx":04D2
            Object.ToolTipText     =   "Thumbs.frx":04F2
         EndProperty
         BeginProperty Panel3 {CB0F173F-9E1F-4365-BF3C-6CC52F8C268B} 
            Version         =   258
            Alignment       =   0
            BorderStyle     =   0
            Content         =   1
            Enabled         =   0   'False
            ForeColor       =   -1
            MinimumWidth    =   100
            PanelData       =   0
            ParseTabs       =   -1  'True
            PreferredWidth  =   70
            RightToLeftText =   0   'False
            Text            =   "Thumbs.frx":0520
            Object.ToolTipText     =   "Thumbs.frx":0548
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
         Text            =   "Thumbs.frx":0576
         Object.ToolTipText     =   "Thumbs.frx":05A0
      EndProperty
   End
   Begin VB.Label lblAnzahl 
      BorderStyle     =   1  'Fest Einfach
      Height          =   372
      Left            =   4200
      TabIndex        =   10
      Top             =   120
      Width           =   732
   End
   Begin VB.Label lblAnzahlThumbnails 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Anzahl Thumbnails:"
      Height          =   372
      Left            =   2280
      TabIndex        =   9
      Top             =   120
      Width           =   1932
   End
End
Attribute VB_Name = "frmThumbs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Der Optionbutton kann zwar Bilder aus Unicode-Verzeichnissen darstellen, weil ich LoadPictureW benutze, aber seine Caption-Eigenschaft
'kann keine Unicode-Dateinamen darstellen
'Deshalb lege ich über den Optionbutton ein Unicode-Label
Option Explicit

    Private Type POINTAPI
        X  As Long
        Y  As Long
    End Type
    
    Private mbActive                As Boolean
    Private mlCurThumb              As Long
    Private Const SRCCOPY           As Long = &HCC0020
    Private Const STRETCH_HALFTONE  As Long = &H4&
    Private Const SW_RESTORE        As Long = &H9&
    
    Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndParent As Long) As Long
    
    Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
    Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lpPt As POINTAPI) As Long
    Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
    Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
    Private Declare Function UnrealizeObject Lib "gdi32" (ByVal hObject As Long) As Long
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Dim KollIndex As Long
    Public Koll As New Collection
    Private gblnSubdirectories As Boolean
    Dim FolderPath As String
    Dim blnAbbrechen As Boolean

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

'    picThumb.Width = 64
'    picThumb.Height = 64
    picThumb.Width = 180
    picThumb.Height = 180

    picThumb.BackColor = vbButtonFace
    picThumb.AutoRedraw = True
    picThumb.Cls
    
    If picSource.Width <= picThumb.Width - 2 And picSource.Height <= picThumb.Height - 2 Then
        fScale = 1
    Else
        fScale = IIf(picSource.Width > picSource.Height, (picThumb.Width - 2) / picSource.Width, (picThumb.Height - 2) / picSource.Height)
    End If
    lWidth = picSource.Width * fScale
    lHeight = picSource.Height * fScale
    lLeft = Int((picThumb.Width - lWidth) / 2)
    lTop = Int((picThumb.Height - lHeight) / 2)
    
    'Store the original ForeColor
    lForeColor = picThumb.ForeColor
    
    'Set picEdit's stretch mode to halftone (this may cause misalignment of the brush)
    lOrigMode = SetStretchBltMode(picThumb.hdc, STRETCH_HALFTONE)
    
    'Realign the brush...
    'Get picEdit's brush by selecting a dummy brush into the DC
    hDummyBrush = CreateSolidBrush(lForeColor)
    hBrush = SelectObject(picThumb.hdc, hDummyBrush)
    'Reset the brush (This will force windows to realign it when it's put back)
    lRet = UnrealizeObject(hBrush)
    'Set picEdit's brush alignment coordinates to the left-top of the bitmap
    lRet = SetBrushOrgEx(picThumb.hdc, lLeft, lTop, uBrushOrigPt)
    'Now put the original brush back into the DC at the new alignment
    hDummyBrush = SelectObject(picThumb.hdc, hBrush)
    
    'Stretch the bitmap
    lRet = StretchBlt(picThumb.hdc, lLeft, lTop, lWidth, lHeight, _
            picSource.hdc, 0, 0, picSource.Width, picSource.Height, SRCCOPY)
    
    'Set the stretch mode back to it's original mode
    lRet = SetStretchBltMode(picThumb.hdc, lOrigMode)
    
    'Reset the original alignment of the brush...
    'Get picEdit's brush by selecting the dummy brush back into the DC
    hBrush = SelectObject(picThumb.hdc, hDummyBrush)
    'Reset the brush (This will force windows to realign it when it's put back)
    lRet = UnrealizeObject(hBrush)
    'Set the brush alignment back to the original coordinates
    lRet = SetBrushOrgEx(picThumb.hdc, uBrushOrigPt.X, uBrushOrigPt.Y, uBrushOrigPt)
    'Now put the original brush back into picEdit's DC at the original alignment
    hDummyBrush = SelectObject(picThumb.hdc, hBrush)
    'Get rid of the dummy brush
    lRet = DeleteObject(hDummyBrush)
    
    'Restore the original ForeColor
    picThumb.ForeColor = lForeColor

    'picThumb.Line (lLeft - 1, lTop - 1)-Step(lWidth + 1, lHeight + 1), &H0&, B
End Sub

Private Sub CreateThumbs()
    Dim iMaxLen As Integer
    Dim X       As Long
    Dim Y       As Long
    Dim lIdx    As Long
    Dim lPicCnt As Long
    Dim lFilCnt As Long
    Dim sText   As String
    Dim TeilFile As String
    Dim CreateThumbsZähler As Long

    Screen.MousePointer = vbHourglass
    picSlide.Move 0, 0, optThumb(0).Width, optThumb(0).Height
    picSlide.Visible = False
    picSlide.BackColor = vbButtonFace
    Set picSlide.Font = optThumb(0).Font
    While optThumb.Count > 1
        Unload optThumb(optThumb.Count - 1)
    Wend
    TimerFrmThumbs.Enabled = True
    DoEvents
    On Error Resume Next
    lFilCnt = Koll.Count
    If Koll.Count > 0 Then
        KollIndex = 1
        blnAbbrechen = False
        CreateThumbsZähler = 0
        For lIdx = 0 To Koll.Count - 1
            StatBarU.Panels(0).Text = Koll.Item(KollIndex)
            Set picLoad.Picture = LoadPicture()
            picLoad.Cls
            Err.Clear
'            If InStr(1, LCase$(Koll.Item(KollIndex)), ".ico") > 0 _
'              Or InStr(1, LCase$(Koll.Item(KollIndex)), ".cur") > 0 Then
'                Set picLoad.Picture = LoadPictureW(Koll.Item(KollIndex), vbLPLargeShell, vbLPDefault)
'            Else
                Set picLoad.Picture = LoadPictureWfrmThumbs(Koll.Item(KollIndex))   'Unicode
'            End If
            Call CreateThumbPic(picLoad, picThumb)
            If lPicCnt > 0 Then                                                     'beim ersten mal ist lPicCnt=0
                Load optThumb(lPicCnt)
                Set optThumb(lPicCnt).Container = picSlide
                Load Ulabel(lPicCnt)
            End If
            TeilFile = file_sep(Koll.Item(KollIndex), True)
            Set optThumb(lPicCnt).Picture = picThumb.image
            sText = TeilFile
            iMaxLen = optThumb(lPicCnt).Width - 15
            If picSlide.TextWidth(sText) > iMaxLen Then
                iMaxLen = iMaxLen - picSlide.TextWidth("...")
            End If
            While picSlide.TextWidth(sText) > iMaxLen
                sText = Left$(sText, Len(sText) - 1)
            Wend
            If iMaxLen < optThumb(lPicCnt).Width - 15 Then
                sText = sText & "..."
            End If
            'optThumb(lPicCnt).Caption = sText
            Ulabel(lPicCnt).Text = sText
            Ulabel(lPicCnt).Tag = Koll.Item(KollIndex)                              'Tag enthält den ganzen Dateiname
            optThumb(lPicCnt).Visible = True
            Ulabel(lPicCnt).Visible = True
            lPicCnt = lPicCnt + 1
            KollIndex = KollIndex + 1
            If blnAbbrechen = True Then
                Exit For
            End If
            CreateThumbsZähler = CreateThumbsZähler + 1
            lblAnzahl.Caption = CreateThumbsZähler
            DoEvents
        Next lIdx
        'Free the unneeded resources
        Set picLoad.Picture = LoadPicture()
        Set picThumb.Picture = LoadPicture()
        'optThumb(0).Value = True
        mlCurThumb = 0
        Call Form_Resize
        picSlide.Visible = True
        TimerFrmThumbs.Enabled = False
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdAbbrechen_Click()
    blnAbbrechen = True
End Sub
'
'Private Sub Form_Load()
'    Dim msg As String
'
'    gblnFrmThumbsLoaded = True
'    Set Koll = Nothing
'    Call KollFüllen
'    If Koll.Count <> 0 Then
'        Me.Show
'        vsbSlide.Visible = False
'        Me.WindowState = vbMaximized
'        SetParent Progbar.hWnd, StatBarU.hWnd
'        StatBarU.Panels(0).PreferredWidth = (Me.Width \ Screen.TwipsPerPixelX) \ 3
'        StatBarU.Panels(1).PreferredWidth = (Me.Width \ Screen.TwipsPerPixelX) \ 3
'        StatBarU_ResizedControlWindow
'        Progbar.Max = Koll.Count
'        CreateThumbs
'    Else
'        msg = "Die Datenbank enthält keine Bilder." & vbNewLine
'        msg = msg & "es konnten keine Thumbnails erzeugt werden."
'        MsgBox msg
'        Unload frmThumbs
'    End If
'End Sub
'
'Private Sub Form_Resize()
'    Dim i As Long
'    Dim X       As Long
'    Dim Y       As Long
'    Dim lIdx    As Long
'    Dim lCols   As Long
'
'    If Koll.Count = 0 Then Exit Sub
'    If Me.WindowState <> vbMinimized Then
'        If Me.Width < 346 * Screen.TwipsPerPixelX Then
'            Me.Width = 346 * Screen.TwipsPerPixelX
'        ElseIf Me.Height < 378 * Screen.TwipsPerPixelY Then
'            Me.Height = 378 * Screen.TwipsPerPixelY
'        Else
'            picFrame.Move 0, cmdAbbrechen.Height + 30, Me.ScaleWidth, Me.ScaleHeight - StatBarU.Height - 70
'            vsbSlide.Move picFrame.ScaleWidth - vsbSlide.Width, 0, vsbSlide.Width, picFrame.ScaleHeight
'            lCols = Int((picFrame.ScaleWidth - vsbSlide.Width) / optThumb(0).Width)
'            For i = 0 To optThumb.Count - 1
'                optThumb(i).BackColor = vbButtonFace
'            Next i
'            'Bisher liegt sowohl die Liste optThumb(index) wie auch die Liste Ulabel(index) ein Element über dem anderen
'
'            For lIdx = 0 To optThumb.Count - 1
'                'jetzt wird jedes Element der Liste optThumb(index) wie auch der Liste Ulabel(index) an seinen eigenen Platz verschoben
'                X = (lIdx Mod lCols) * optThumb(0).Width
'                Y = Int(lIdx / lCols) * optThumb(0).Height
'                optThumb(lIdx).Move X, Y
'                Ulabel(lIdx).Move X + 10, Y + 10
'                Ulabel(lIdx).ZOrder
'            Next lIdx
'
'            picSlide.Width = lCols * optThumb(0).Width
'            picSlide.Height = Int(optThumb.Count / lCols) * optThumb(0).Height
'            If Int(optThumb.Count / lCols) < (optThumb.Count / lCols) Then
'                picSlide.Height = picSlide.Height + optThumb(0).Height
'            End If
'            vsbSlide.Value = 0
'            vsbSlide.Max = picSlide.Height - picFrame.ScaleHeight
'            If vsbSlide.Max < 0 Then
'                vsbSlide.Max = 0
'                vsbSlide.Enabled = False
'            Else
'                vsbSlide.Enabled = True
'                vsbSlide.SmallChange = optThumb(0).Height
'                vsbSlide.LargeChange = picFrame.ScaleHeight
'            End If
'        End If
'    End If
'    'cmdAbbrechen.Visible = False
'    vsbSlide.Visible = True
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'    Dim lIdx As Long
'
'    For lIdx = 1 To optThumb.Count - 1
'        Unload optThumb(lIdx)
'    Next lIdx
'    gblnFrmThumbsLoaded = False
'End Sub

Private Sub optThumb_Click(Index As Integer)
    Dim i As Long
    
    For i = 0 To optThumb.Count - 1
        optThumb(i).BackColor = vbButtonFace
    Next i
    optThumb(Index).BackColor = vbBlue
End Sub

Private Sub optThumb_DblClick(Index As Integer)
    Dim txtSuchen As String
    Dim strFind As String
    Dim msg As String
    
    'MsgBox "Ulabel(Index)=" & Ulabel(Index) & "optThumb(Index).Tag" & optThumb(Index).Tag
    Me.Hide
    gblnComeFromThumbs = True
    gstrFRODN = Ulabel(Index).Tag
    Call Form1.BildAnzeigen
    gblnComeFromThumbs = False
    'In der DbGridForm soll derselbe Satz eingestellt werden, wie der dem geklickten Bild entsprechende
    ' Der Dateiname muss verwandelt werden in die Form wie er in der Datenbank steht
    txtSuchen = Replace(gstrFRODN, gstrFotosMdbLocation & "\", "+:\")
    'strFind = "Dateiname like '*" & txtSuchen & "*'"
    strFind = LoadResString(1028 + Sprache) & " like '*" & txtSuchen & "*'"
    On Error GoTo 0
    'hier wird zweimal DbGridForm_RowColChange aufgerufen, bei MoveFirst und bei Find und in der Folge die Function Form1.Bildanzeigen
    'das will ich verhindern weil es flackert
    gblnComeFromThumbs = True
    DbGridForm.Adodc1.Recordset.MoveFirst                   'unbedingt vor ...Find musst Du .MoveFirst machen
    DbGridForm.Adodc1.Recordset.Find strFind                'sonst kommt error 3021 wenn der Satz weiter vorn steht als der aktuelle
    gblnComeFromThumbs = False
    If DbGridForm.DBGridNeu.SelBookmarks.Count = 1 Then
        DbGridForm.DBGridNeu.SelBookmarks.Remove 0
    End If
    DbGridForm.DBGridNeu.SelBookmarks.Add DbGridForm.rsDataGrid.Bookmark
End Sub


Private Sub TimerFrmThumbs_Timer()
  On Error Resume Next
  Progbar.Value = KollIndex
  On Error GoTo 0
End Sub

Private Sub vsbSlide_Change()
    picSlide.Top = -vsbSlide.Value
    On Error Resume Next
    picFrame.SetFocus
    On Error GoTo 0
End Sub

Private Sub vsbSlide_Scroll()
    vsbSlide_Change
End Sub

Private Sub KollFüllen()
    Dim MyRecordset As adodb.Recordset
    Dim Filename As String
    Dim DateinamenErweiterung As String
    
    Screen.MousePointer = vbHourglass
    gblnSubdirectories = False
    'In GERBING Fotoalbum wird anstelle von Function Rekursive der von der Abfrage gelieferte recordset benutzt
    'Call Rekursive(FolderPath, "*.*")
    
    ' Recordset erstellen und öffnen adOpenStatic
    Set MyRecordset = New adodb.Recordset
    With MyRecordset
        If gblnWasHeadClick = False Then
            .Source = DbGridForm.Adodc1.RecordSource
        Else
            .Source = DbGridForm.SQLneuHeadClick
        End If
        .ActiveConnection = DBsql
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    While Not MyRecordset.EOF
        Filename = Replace(MyRecordset.Fields(LoadResString(1028 + Sprache)), "+:\", gstrFotosMdbLocation & "\")  '1028=Dateiname  'Gerbing 07.11.2011
        '---------------------------------------------------
        DateinamenErweiterung = UCase(Right(Filename, 3))
        Select Case DateinamenErweiterung
            Case "BMP", "DIB", "GIF", "JPG", "TIF", "PNG"
            'Case "BMP", "DIB", "GIF", "JPG"
                'Koll.Add txtOrdnerMitBildern & "\" & Dateiname                      'Gerbing 20.01.2009
                Koll.Add Filename
        End Select
        '---------------------------------------------------
        MyRecordset.MoveNext
    Wend
    MyRecordset.Close
    Screen.MousePointer = vbNormal                                                          'Gerbing 07.01.2009
End Sub

Private Sub StatBarU_ResizedControlWindow()
  Dim x1 As Long
  Dim x2 As Long
  Dim y1 As Long
  Dim y2 As Long

  StatBarU.Panels(1).GetRectangle x1, y1, x2, y2
  MoveWindow Progbar.hWnd, x1, y1 + 1, x2 - x1 - 4, y2 - y1 - 2, 1
End Sub

