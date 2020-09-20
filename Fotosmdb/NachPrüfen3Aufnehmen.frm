VERSION 5.00
Object = "{FCCB83BF-E483-4317-9FF2-A460758238B5}#1.5#0"; "CBLCtlsU.ocx"
Begin VB.Form NachPrüfen3Aufnehmen 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Prüfen3 - Gefundene Dateien in die Datenbank aufnehmen"
   ClientHeight    =   8004
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   12840
   Icon            =   "NachPrüfen3Aufnehmen.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8004
   ScaleWidth      =   12840
   StartUpPosition =   1  'Fenstermitte
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   375
      Left            =   8520
      ScaleHeight     =   324
      ScaleWidth      =   324
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton btnAlleMarkieren 
      Caption         =   "Alle mar&kieren"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   7440
      Width           =   3492
   End
   Begin VB.CommandButton btnAufnehmen 
      Caption         =   "&markierte Dateien aufnehmen"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Zum Markieren können Sie die Tasten Umsch und Strg zu Hilfe nehmen"
      Top             =   7440
      Width           =   3492
   End
   Begin VB.CommandButton btnAbbrechen 
      Caption         =   "&Abbrechen"
      Height          =   375
      Left            =   9240
      TabIndex        =   0
      Top             =   7440
      Width           =   3492
   End
   Begin CBLCtlsLibUCtl.ListBox lstZusätzlicheDateien 
      Height          =   7092
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   8052
      _cx             =   14203
      _cy             =   12509
      AllowDragDrop   =   0   'False
      AllowItemSelection=   -1  'True
      AlwaysShowVerticalScrollBar=   -1  'True
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   0
      ColumnWidth     =   -1
      DisabledEvents  =   1048808
      DontRedraw      =   0   'False
      DragScrollTimeBase=   -1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      HasStrings      =   -1  'True
      HoverTime       =   -1
      IMEMode         =   -1
      InsertMarkColor =   0
      InsertMarkStyle =   1
      IntegralHeight  =   0   'False
      ItemHeight      =   -1
      Locale          =   1024
      MousePointer    =   0
      MultiColumn     =   0   'False
      MultiSelect     =   1
      OLEDragImageStyle=   0
      OwnerDrawItems  =   0
      ProcessContextMenuKeys=   -1  'True
      ProcessTabs     =   -1  'True
      RegisterForOLEDragDrop=   0   'False
      RightToLeft     =   0
      ScrollableWidth =   1500
      Sorted          =   0   'False
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      ToolTips        =   0
      UseSystemFont   =   0   'False
      VirtualMode     =   0   'False
   End
   Begin VB.Image Image1 
      Height          =   372
      Left            =   9000
      Top             =   1200
      Width           =   372
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sie erhalten ein Vorschaubild, wenn Sie auf den Dateiname rechtsklicken"
      Height          =   972
      Left            =   8280
      TabIndex        =   3
      Top             =   120
      Width           =   4452
   End
End
Attribute VB_Name = "NachPrüfen3Aufnehmen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim EsGibtAufzunehmende As Boolean
    Dim MyAppPath As String
    Public KollZusätzlicheDateien As New Collection
    
    '----------------------------------------------ab hier für Video Thumbnails-------Gerbing 06.04.2017
    'From Windows SDK header file propkey.h:
    Private Const SCID_THUMBNAILSTREAM As String = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9},27"
    'Requires Windows XP SP2 or later:
    Private Declare Function PropVariantToVariant Lib "propsys" ( _
        ByRef PropVar As Any, _
        ByRef Var As Variant) As Long
    Private ShellObject As shell32.Shell
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
    
Private Sub btnAbbrechen_Click()
    Do Until NachPrüfen3Löschen.KollZusätzlicheDateien.Count = 0                    'Gerbing 03.11.2013
        NachPrüfen3Löschen.KollZusätzlicheDateien.Remove 1
    Loop
    Do Until NachPrüfen3Aufnehmen.KollZusätzlicheDateien.Count = 0                  'Gerbing 03.11.2013
        NachPrüfen3Aufnehmen.KollZusätzlicheDateien.Remove 1
    Loop
    Me.Hide
End Sub

Private Sub btnAlleMarkieren_Click()
    Dim n As Long
    
    lstZusätzlicheDateien.Visible = False                                                   'Gerbing 26.10.2013
    For n = 0 To lstZusätzlicheDateien.ListItems.Count - 1
        'lstZusätzlicheDateien.Selected(n) = True
        lstZusätzlicheDateien.ListItems(n).Selected = True
    Next n
    lstZusätzlicheDateien.Visible = True                                                    'Gerbing 26.10.2013
End Sub

Private Sub btnAufnehmen_Click()
    Dim i As Long
    Dim j As Long
    Dim n As Long
    
    i = 0
    j = 1
    Screen.MousePointer = vbHourglass
    Me.MousePointer = vbHourglass
    For i = 0 To lstZusätzlicheDateien.ListItems.Count - 1                                      'Gerbing 26.10.2013
        If lstZusätzlicheDateien.ListItems(i).Selected = False Then
            KollZusätzlicheDateien.Remove j
        Else
            j = j + 1
        End If
        DoEvents
    Next i
    For i = 1 To KollZusätzlicheDateien.Count
        NeueDatensätzeGenerieren.List1.ListItems.Add KollZusätzlicheDateien.Item(i)
        EsGibtAufzunehmende = True
        DoEvents
    Next i
    
    lstZusätzlicheDateien.Visible = True
    Screen.MousePointer = vbDefault
    Me.MousePointer = vbDefault
    Form1.FehlerGefunden = False
    Form1.txtFehlerU.Text = ""
    Me.Hide
    LogLesen.Hide
    If EsGibtAufzunehmende = True Then
        Call Form1.btnGenerieren_Click
        Call Form1.btnReset_Click                       'Gerbing 04.02.2008
    End If
    lstZusätzlicheDateien.ListItems.RemoveAll
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    Call AnpassenNutzerWunsch(Me)                                       'Gerbing 11.03.2017
    Me.Caption = LoadResString(1341 + Sprache)      'Prüfen3 - Gefundene Dateien in die Datenbank aufnehmen
    btnAufnehmen.Caption = LoadResString(1342 + Sprache)        '&markierte Dateien aufnehmen
    btnAbbrechen.Caption = LoadResString(3013 + Sprache)        '&Abbrechen
    btnAufnehmen.ToolTipText = LoadResString(1432 + Sprache)    'Zum Markieren können Sie die Tasten Umsch und Strg zu Hilfe nehmen
    btnAlleMarkieren.Caption = LoadResString(1518 + Sprache) 'Alle mar&kieren
    Label1.Caption = LoadResString(1522 + Sprache)  'Sie erhalten ein Vorschaubild, wenn Sie auf den Dateiname rechtsklicken
    
    'lstZusätzlicheDateien.MultiSelect = 2 muss in der Entwicklungsumgebung eingestellt werden
    EsGibtAufzunehmende = False
    If gblnSQLServerVersion = True Then
        MyAppPath = PublicLocationFotos
    Else
        MyAppPath = AppPath
    End If
    
'    Set ShellObject = New shell32.Shell                                 'Gerbing 06.04.2017
'    Set GdipTool = New GdipTool                                         'Gerbing 06.04.2017
    
    'Set ShellObject = New shell32.Shell                        'Gerbing 09.08.2017
    On Error Resume Next
    'Set ShellObject = CreateObject("Shell.Application")
    Set ShellObject = CreateObject(CVar("Shell.Application"))
    Set GdipTool = New GdipTool                                 'Gerbing 09.08.2017
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lstZusätzlicheDateien.width = NachPrüfen3Aufnehmen.width - 4600            'Gerbing 22.11.2006
    lstZusätzlicheDateien.height = NachPrüfen3Aufnehmen.height - 1240
    btnAufnehmen.Top = Me.height - 975
    btnAbbrechen.Top = Me.height - 975
    btnAlleMarkieren.Top = Me.height - 975
    Label1.Left = lstZusätzlicheDateien.width + 200
    Image1.Left = lstZusätzlicheDateien.width + 200
    Picture1.Left = lstZusätzlicheDateien.width + 200
    On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    EsGibtAufzunehmende = False
    Cancel = True
    Me.Hide
End Sub

Public Sub BildAnzeigen(EchterStandort)
    Dim Bildbreite As Long
    Dim Bildhöhe As Long
    Dim Image1Top As Long
    Dim Image1Left As Long
    Dim BHV As Double   'BreitenHöhenVerhältnis
    Dim strTemp As String

    strTemp = Replace(EchterStandort, "+:\", MyAppPath & "\")        'Gerbing 11.04.2005
    Bildbreite = 4000
    Bildhöhe = 3000
    Image1Top = Picture1.Top
    Image1Left = Picture1.Left

    On Error GoTo 0
    ERR = 0
    Screen.MousePointer = vbHourglass
    Image1.Visible = False
    Picture1.Visible = False
    Picture1.AutoSize = True
    On Error Resume Next
    'Picture1.Picture = LoadPicture(strTemp)
    'Call LoadPictureW(strTemp, Me)
    'Set Picture1.Picture = CreateThumbnailFromFile(strTemp, 100)       'Gerbing 30.09.2015
    Set Picture1.Picture = CreateStdPictureFromFile(strTemp)            'Gerbing 02.10.2015
    If ERR.Number <> 0 Then
        Screen.MousePointer = vbDefault
        Call BildFehler(strTemp)
        Exit Sub
    End If
    On Error GoTo 0
    '-----------------------------------------------------------------------------------------------
    'Untersuchung, ob das Bild größer ist als die Bildbreite/höhe und dessen Konsequenzen
    BHV = Picture1.width / Picture1.height

    If Picture1.width > Bildbreite Or Picture1.height > Bildhöhe Then
    'wenn das Bild größer ist als Bildbreite/Bildhöhe wird es verkleinert
        Image1.Stretch = True
        Image1.Picture = Picture1.Picture
        Select Case BHV
            Case 1.33 To 1.34
                'das Breitenverhältnis ist 4/3 = 1.33
                Image1.Top = Image1Top
                Image1.Left = Image1Left
                Image1.width = Bildbreite
                Image1.height = Bildhöhe
            Case Is < 1.33
                'das Bild ist zu hoch und zu schmal
                Image1.Top = Image1Top
                Image1.Left = Image1Left
                Image1.height = Bildhöhe
                Image1.width = Bildhöhe * BHV
            Case Else
                'das Bild ist zu niedrig und zu breit
                Image1.Top = Image1Top
                Image1.Left = Image1Left
                Image1.width = Bildbreite
                Image1.height = Bildbreite / BHV
        End Select
    Else
    'wenn das Bild nicht größer ist als Bildbreite/Bildhöhe
        'Bild in links oben im Bildbereich anordnen
        Image1.Stretch = True
        Image1.Picture = Picture1.Picture
        Image1.Top = Image1Top
        Image1.Left = Image1Left
        Image1.width = Picture1.width
        Image1.height = Picture1.height
    End If
    Image1.Visible = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub BildFehler(EchterStandort)
    Dim Msg As String
    
'    msg = "Bild kann nicht geladen werden" & NL
'    msg = msg & EchterStandort & NL
'    msg = msg & "Prüfen Sie ob diese Datei existiert" & NL
'    msg = msg & "oder ob es sich um einen verbotenen Dateityp handelt."
    Msg = LoadResString(2056 + Sprache) & vbNewLine
    Msg = Msg & EchterStandort & vbNewLine
    Msg = Msg & LoadResString(2301 + Sprache) & vbNewLine
    Msg = Msg & LoadResString(2302 + Sprache)
    MsgBox Msg, vbInformation
End Sub

Private Sub lstZusätzlicheDateien_MouseUp(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
    Dim folder As shell32.folder                                                'Gerbing 06.04.2017
    Dim ShellFolderItem As shell32.ShellFolderItem                              'Gerbing 06.04.2017
    Dim GdipLoader As GdipLoader                                                'Gerbing 06.04.2017
    Dim PropLong As Variant                                                     'Gerbing 06.04.2017
    Dim PropStream As IUnknown                                                  'Gerbing 06.04.2017
    Dim outname As String                                                       'Gerbing 06.04.2017
    Dim DateinamenErweiterung As String                                         'Gerbing 06.04.2017
    Dim strPath As String                                                       'Gerbing 06.04.2017
    Dim strFile As String                                                       'Gerbing 06.04.2017
    Dim Fotodatei As String                                                     'Gerbing 06.04.2017
    Dim Msg As String                                                           'Gerbing 22.11.2018
    
    If Button = vbRightButton Then
        'Prüfe, ob die DateinamenErweiterung zu einem Bild gehört               'Gerbing 11.04.2005
        'wenn ja Vorschaubild zeigen
        If Not listItem Is Nothing Then                                         'Gerbing 28.05.2019
            Fotodatei = Replace(listItem, "+:\", AppPath & "\")
            Call file_split(Fotodatei, strPath, strFile, DateinamenErweiterung)
            DateinamenErweiterung = UCase(DateinamenErweiterung)
            On Error Resume Next                                                    'Gerbing 14.02.2018
            Select Case DateinamenErweiterung
                Case "BMP", "CUR", "DIB", "EMF", "GIF", "ICO", "JPG", "WMF"         'Gerbing 09.03.2005
                    'nur wenn es tatsächlich eine Bilddatei ist
                    'Call BildAnzeigen(lstZusätzlicheDateien.Text)
                    Call BildAnzeigen(listItem)
                Case "AVI", "MPG", "MOV", "WMV", "ASF", "MP4", "ASX", "MKV", "FLV"  'Gerbing 10.12.2017 bei Videos
                    Call BildAnzeigen("")
                    Set folder = ShellObject.NameSpace(strPath)
                    If folder Is Nothing Then
                        Msg = strPath & " folder Is wrong"
                        'MsgBox "Folder Is Nothing"
                        MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotosmdb"), vbInformation   'Gerbing 22.11.2018
                    Else
                        Set GdipLoader = New GdipLoader
                    End If
                    Set ShellFolderItem = folder.ParseName(strFile & "." & DateinamenErweiterung)
                    PropLong = ShellFolderItem.ExtendedProperty(SCID_PerceivedType)
                    If Not IsEmpty(PropLong) Then
                        If PropLong = PERCEIVED_TYPE_VIDEO Then
                            On Error Resume Next
                            Set PropStream = ShellFolderItem.ExtendedProperty(SCID_PropStream)
                            If ERR.Number = 0 Then
                                If Not PropStream Is Nothing Then
                                    outname = AppPath & "\TempThumbs\" & ShellFolderItem.Name & ".jpg"
                                    On Error Resume Next
                                    If file_path_exist(strPath & "TempThumbs\") = False Then
                                        MkDir AppPath & "\TempThumbs\"
                                    End If
                                    On Error GoTo 0
                                    GdipTool.PropStream2PicFileScaled PropStream, outname, PFF_JPEG, , 400, 320
                                End If
                            Else
                                Exit Sub
                            End If
                        End If
                        On Error GoTo 0
                    End If
                    Call BildAnzeigen(outname)
                Case Else
                    Call BildAnzeigen("")
            End Select
        End If                                                                      'Gerbing 28.05.2019
    End If
End Sub
