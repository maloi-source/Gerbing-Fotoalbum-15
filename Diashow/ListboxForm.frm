VERSION 5.00
Object = "{9FC6639B-4237-4FB5-93B8-24049D39DF74}#1.5#0"; "ExLvwU.ocx"
Object = "{A8F9B8E7-E699-4FCE-A647-72C877F8E632}#1.8#0"; "editctlsu.ocx"
Begin VB.Form ListBoxForm 
   BackColor       =   &H00C0C0C0&
   Caption         =   "txtIPTCInfo"
   ClientHeight    =   6684
   ClientLeft      =   2376
   ClientTop       =   2388
   ClientWidth     =   13548
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ListboxForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6684
   ScaleWidth      =   13548
   Begin VB.CommandButton btnAlleHäkchenEntfernen 
      Height          =   372
      Left            =   720
      Picture         =   "ListboxForm.frx":000C
      Style           =   1  'Grafisch
      TabIndex        =   6
      ToolTipText     =   "alle Markierungen entfernen"
      Top             =   120
      Width           =   372
   End
   Begin VB.CommandButton btnAlleHäkchenSetzen 
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
      Picture         =   "ListboxForm.frx":0596
      Style           =   1  'Grafisch
      TabIndex        =   5
      ToolTipText     =   "alle markieren"
      Top             =   120
      Width           =   372
   End
   Begin VB.Frame FrameExifIptc 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   8280
      TabIndex        =   1
      Top             =   0
      Width           =   5172
      Begin VB.CheckBox chkIptcAnzeigen 
         BackColor       =   &H00C0C0C0&
         Caption         =   "IPTC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Setzen Sie hier ein Häkchen, wenn Sie die IPTC-Felder des markierten Satzes im Drag&Drop-Container sehen wollen"
         Top             =   360
         Width           =   4812
      End
      Begin VB.CheckBox chkExifAnzeigen 
         BackColor       =   &H00C0C0C0&
         Caption         =   "EXIF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Setzen Sie hier ein Häkchen, wenn Sie die EXIF-Felder des markierten Satzes im Drag&Drop-Container sehen wollen"
         Top             =   120
         Width           =   4812
      End
      Begin EditCtlsLibUCtl.TextBox txtEXIFInfo 
         Height          =   3012
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   4932
         _cx             =   8700
         _cy             =   5313
         AcceptNumbersOnly=   0   'False
         AcceptTabKey    =   0   'False
         AllowDragDrop   =   -1  'True
         AlwaysShowSelection=   0   'False
         Appearance      =   1
         AutoScrolling   =   0
         BackColor       =   -2147483643
         BorderStyle     =   0
         CancelIMECompositionOnSetFocus=   0   'False
         CharacterConversion=   0
         CompleteIMECompositionOnKillFocus=   0   'False
         DisabledBackColor=   -1
         DisabledEvents  =   3075
         DisabledForeColor=   -1
         DisplayCueBannerOnFocus=   0   'False
         DontRedraw      =   0   'False
         DoOEMConversion =   0   'False
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
         FormattingRectangleHeight=   0
         FormattingRectangleLeft=   0
         FormattingRectangleTop=   0
         FormattingRectangleWidth=   0
         HAlignment      =   0
         HoverTime       =   -1
         IMEMode         =   -1
         InsertMarkColor =   0
         InsertSoftLineBreaks=   0   'False
         LeftMargin      =   -1
         MaxTextLength   =   -1
         Modified        =   0   'False
         MousePointer    =   0
         MultiLine       =   -1  'True
         OLEDragImageStyle=   0
         PasswordChar    =   0
         ProcessContextMenuKeys=   -1  'True
         ReadOnly        =   0   'False
         RegisterForOLEDragDrop=   0   'False
         RightMargin     =   -1
         RightToLeft     =   0
         ScrollBars      =   3
         SelectedTextMousePointer=   0
         SupportOLEDragImages=   -1  'True
         TabWidth        =   -1
         UseCustomFormattingRectangle=   0   'False
         UsePasswordChar =   0   'False
         UseSystemFont   =   0   'False
         CueBanner       =   "ListboxForm.frx":0B20
         Text            =   "ListboxForm.frx":0B40
      End
      Begin EditCtlsLibUCtl.TextBox txtIPTCInfo 
         Height          =   3012
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   4932
         _cx             =   8700
         _cy             =   5313
         AcceptNumbersOnly=   0   'False
         AcceptTabKey    =   0   'False
         AllowDragDrop   =   -1  'True
         AlwaysShowSelection=   0   'False
         Appearance      =   1
         AutoScrolling   =   0
         BackColor       =   -2147483643
         BorderStyle     =   0
         CancelIMECompositionOnSetFocus=   0   'False
         CharacterConversion=   0
         CompleteIMECompositionOnKillFocus=   0   'False
         DisabledBackColor=   -1
         DisabledEvents  =   3075
         DisabledForeColor=   -1
         DisplayCueBannerOnFocus=   0   'False
         DontRedraw      =   0   'False
         DoOEMConversion =   0   'False
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
         FormattingRectangleHeight=   0
         FormattingRectangleLeft=   0
         FormattingRectangleTop=   0
         FormattingRectangleWidth=   0
         HAlignment      =   0
         HoverTime       =   -1
         IMEMode         =   -1
         InsertMarkColor =   0
         InsertSoftLineBreaks=   0   'False
         LeftMargin      =   -1
         MaxTextLength   =   -1
         Modified        =   0   'False
         MousePointer    =   0
         MultiLine       =   -1  'True
         OLEDragImageStyle=   0
         PasswordChar    =   0
         ProcessContextMenuKeys=   -1  'True
         ReadOnly        =   0   'False
         RegisterForOLEDragDrop=   0   'False
         RightMargin     =   -1
         RightToLeft     =   0
         ScrollBars      =   3
         SelectedTextMousePointer=   0
         SupportOLEDragImages=   -1  'True
         TabWidth        =   -1
         UseCustomFormattingRectangle=   0   'False
         UsePasswordChar =   0   'False
         UseSystemFont   =   0   'False
         CueBanner       =   "ListboxForm.frx":0B60
         Text            =   "ListboxForm.frx":0B80
      End
   End
   Begin VB.CommandButton btnAktionAusführen 
      Caption         =   "&Aktion wählen und ausführen..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin ExLVwLibUCtl.ExplorerListView ExLVwU 
      Height          =   3252
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   3612
      _cx             =   6371
      _cy             =   5736
      AbsoluteBkImagePosition=   0   'False
      AllowHeaderDragDrop=   -1  'True
      AllowLabelEditing=   -1  'True
      AlwaysShowSelection=   -1  'True
      Appearance      =   1
      AutoArrangeItems=   0
      AutoSizeColumns =   -1  'True
      BackColor       =   -2147483643
      BackgroundDrawMode=   0
      BkImagePositionX=   0
      BkImagePositionY=   0
      BkImageStyle    =   2
      BlendSelectionLasso=   -1  'True
      BorderSelect    =   0   'False
      BorderStyle     =   0
      CallBackMask    =   0
      CheckItemOnSelect=   0   'False
      ClickableColumnHeaders=   -1  'True
      ColumnHeaderVisibility=   2
      DisabledEvents  =   3145725
      DontRedraw      =   0   'False
      DragScrollTimeBase=   -1
      DrawImagesAsynchronously=   0   'False
      EditBackColor   =   -2147483643
      EditForeColor   =   -2147483640
      EditHoverTime   =   -1
      EditIMEMode     =   -1
      EmptyMarkupTextAlignment=   1
      Enabled         =   -1  'True
      FilterChangedTimeout=   -1
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
      FullRowSelect   =   2
      GridLines       =   -1  'True
      GroupFooterForeColor=   -2147483640
      GroupHeaderForeColor=   -2147483640
      GroupMarginBottom=   0
      GroupMarginLeft =   0
      GroupMarginRight=   0
      GroupMarginTop  =   12
      GroupSortOrder  =   0
      HeaderFullDragging=   -1  'True
      HeaderHotTracking=   0   'False
      HeaderHoverTime =   -1
      HeaderOLEDragImageStyle=   0
      HideLabels      =   0   'False
      HotForeColor    =   -1
      HotMousePointer =   0
      HotTracking     =   0   'False
      HotTrackingHoverTime=   -1
      HoverTime       =   -1
      IMEMode         =   -1
      IncludeHeaderInTabOrder=   0   'False
      InsertMarkColor =   0
      ItemActivationMode=   2
      ItemAlignment   =   1
      ItemBoundingBoxDefinition=   70
      ItemHeight      =   17
      JustifyIconColumns=   0   'False
      LabelWrap       =   -1  'True
      MinItemRowsVisibleInGroups=   0
      MousePointer    =   0
      MultiSelect     =   0   'False
      OLEDragImageStyle=   0
      OutlineColor    =   -2147483633
      OwnerDrawn      =   0   'False
      ProcessContextMenuKeys=   -1  'True
      Regional        =   0   'False
      RegisterForOLEDragDrop=   0   'False
      ResizableColumns=   -1  'True
      RightToLeft     =   0
      ScrollBars      =   1
      SelectedColumnBackColor=   -1
      ShowFilterBar   =   0   'False
      ShowGroups      =   0   'False
      ShowHeaderChevron=   0   'False
      ShowHeaderStateImages=   0   'False
      ShowStateImages =   0   'False
      ShowSubItemImages=   0   'False
      SimpleSelect    =   0   'False
      SingleRow       =   -1  'True
      SnapToGrid      =   0   'False
      SortOrder       =   0
      SupportOLEDragImages=   -1  'True
      TextBackColor   =   -1
      TileViewItemLines=   1
      TileViewLabelMarginBottom=   0
      TileViewLabelMarginLeft=   0
      TileViewLabelMarginRight=   0
      TileViewLabelMarginTop=   0
      TileViewSubItemForeColor=   -1
      TileViewTileHeight=   -1
      TileViewTileWidth=   -1
      ToolTips        =   3
      UnderlinedItems =   0
      UseMinColumnWidths=   0   'False
      UseSystemFont   =   0   'False
      UseWorkAreas    =   0   'False
      View            =   2
      VirtualMode     =   0   'False
      EmptyMarkupText =   "ListboxForm.frx":0BA0
      FooterIntroText =   "ListboxForm.frx":0BC0
   End
End
Attribute VB_Name = "ListBoxForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'txtEXIFInfo ist jetzt ein Timosoft Control unicode fähig                           'Gerbing 13.11.2015
Option Explicit

    Implements ISubclassedWindow

    Public ExLVwUIndex As Long
    Public GPSLatitude As String                                                    'Gerbing 27.08.2015
    Public GPSLatitudeRef As String
    Public GPSLongitude As String
    Public GPSLongitudeRef As String
        
    'Gerbing 04.07.2019
    Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
    Private Const LOCALE_SDECIMAL = &HE                 '  decimal separator

Private Sub Command1_Click()
End Sub

Private Sub btnAlleHäkchenEntfernen_Click()                                         'Gerbing 22.03.2016
    Dim i As Long
    
    For i = 0 To ExLVwU.ListItems.Count - 1
        ExLVwU.ListItems(i).StateImageIndex = 1
        'ExLVwU.ListItems(i).StateImageIndex = 3
    Next
End Sub

Private Sub btnAlleHäkchenSetzen_Click()                                            'Gerbing 22.03.2016
    Dim i As Long
    
    For i = 0 To ExLVwU.ListItems.Count - 1
        ExLVwU.ListItems(i).StateImageIndex = 2
        'ExLVwU.ListItems(i).StateImageIndex = 4
    Next
End Sub

Private Sub ExLVwU_Click(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
    On Error Resume Next                                                            'Gerbing 11.10.2016
    DiashowForm.blnListBoxNeuDblClick = True                                        'Gerbing 21.03.2016
    gblnListBoxNeuDblClick = True                                                   'Gerbing 21.03.2016
    ExLVwUIndex = listItem.ItemData                                                 'Gerbing 21.03.2016
    gblStrAktuellGezeigtesBild = listItem                                           'Gerbing 21.03.2016
    Call DiashowForm.BildAnzeigen                                                   'Gerbing 21.03.2016
    On Error GoTo 0                                                                 'Gerbing 11.10.2016
    Exit Sub                                                                        'Gerbing 21.03.2016
End Sub

Private Sub ExLVwU_ItemStateImageChanging(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal previousStateImageIndex As Long, newStateImageIndex As Long, ByVal causedBy As ExLVwLibUCtl.StateImageChangeCausedByConstants, cancelChange As Boolean)
    '1 und 3 heißt unchecked
    '2 und 4 heißt checked
    'das wird festgelegt in ListBoxForm.Form_Initialize aus dem Ordner AppPath & "\res\"
    If previousStateImageIndex = 1 And newStateImageIndex = 2 Then
        newStateImageIndex = 4
        Exit Sub
    End If
    If previousStateImageIndex = 2 And newStateImageIndex = 3 Then
        newStateImageIndex = 1
        Exit Sub
    End If
    If previousStateImageIndex = 4 And newStateImageIndex = 1 Then
        newStateImageIndex = 3
        Exit Sub
    End If
    If previousStateImageIndex = 3 And newStateImageIndex = 4 Then
        newStateImageIndex = 1
        Exit Sub
    End If
End Sub

Private Sub Form_Initialize()
    Const ILC_COLOR24 = &H18
    Const ILC_COLOR32 = &H20
    Const ILC_MASK = &H1
    Const IMAGE_ICON = 1
    Const LR_DEFAULTSIZE = &H40
    Const LR_LOADFROMFILE = &H10
    Dim DLLVerData As DLLVERSIONINFO
    Dim hIcon As Long
    Dim hMod As Long
    Dim iconsDir As String
    Dim iconPath As String
    Dim wfd As WIN32_FIND_DATA
    Dim hSearch As Long
    Dim Cont As Long
    Dim FileName As String

    InitCommonControls                                      'Gerbing 04.03.2013
    
    With DLLVerData
        .cbSize = LenB(DLLVerData)
        DllGetVersion_comctl32 DLLVerData
        bComctl32Version600OrNewer = (.dwMajor >= 6)
    End With

    
    hImgLst = ImageList_Create(16, 16, IIf(bComctl32Version600OrNewer, ILC_COLOR32, ILC_COLOR24) Or ILC_MASK, 14, 0)
    hStateImgLst = ImageList_Create(16, 16, IIf(bComctl32Version600OrNewer, ILC_COLOR32, ILC_COLOR24) Or ILC_MASK, 4, 0)
    ImageList_SetImageCount hStateImgLst, 4
    If Right$(AppPath, 3) = "bin" Then
      iconsDir = AppPath & "\..\res\"
    Else
      iconsDir = AppPath & "\res\"
    End If
    hSearch = FindFirstFileW(StrPtr(iconsDir & "*.ico"), VarPtr(wfd))
    Cont = True
    If hSearch <> INVALID_HANDLE_VALUE Then
        While Cont
            'FileName = StripNulls(StrConv(wfd.cFileName, vbFromUnicode))
            FileName = RemoveNulls(wfd.cFileName)
            FileName = LCase(FileName)
            If (FileName <> ".") And (FileName <> "..") Then
                hIcon = LoadImage(0, StrPtr(iconsDir & FileName), IMAGE_ICON, 16, 16, LR_LOADFROMFILE Or LR_DEFAULTSIZE)
                If hIcon Then
                    Select Case LCase$(FileName)
                        Case "unchecked.ico"
                            ImageList_ReplaceIcon hStateImgLst, 0, hIcon
                        Case "checked.ico"
                            ImageList_ReplaceIcon hStateImgLst, 1, hIcon
                        Case "unchecked1.ico"
                            ImageList_ReplaceIcon hStateImgLst, 2, hIcon
                        Case "checked1.ico"
                            ImageList_ReplaceIcon hStateImgLst, 3, hIcon
                        Case Else
                            ImageList_AddIcon hImgLst, hIcon
                    End Select
                    DestroyIcon hIcon
                End If
            End If
            Cont = FindNextFileW(hSearch, VarPtr(wfd)) ' Get next file
        Wend
    Else
        'MsgBox "Es wurde kein res-Ordner gefunden. Wiederholen Sie die Installation von Diashow" 'Gerbing 20.10.2017
        'MsgBox "Wenn Sie Diashow.exe mehrfach starten, können Sie keine Häkchen im Dateinamen-Fenster setzen" 'Gerbing 21.11.2017
        MsgBox LoadResString(2362 + Sprache)
    End If
End Sub

Private Function ISubclassedWindow_HandleMessage(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal eSubclassID As EnumSubclassID, bCallDefProc As Boolean) As Long
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case eSubclassID
    Case EnumSubclassID.escidFrmMain
      lRet = HandleMessage_Form(hwnd, uMsg, wParam, lParam, bCallDefProc)
    Case Else
      Debug.Print "frmMain.ISubclassedWindow_HandleMessage: Unknown Subclassing ID " & CStr(eSubclassID)
  End Select

StdHandler_Ende:
  ISubclassedWindow_HandleMessage = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in frmMain.ISubclassedWindow_HandleMessage (SubclassID=" & CStr(eSubclassID) & ": ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function

Private Function HandleMessage_Form(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bCallDefProc As Boolean) As Long
  Const WM_NOTIFYFORMAT = &H55
  Const WM_USER = &H400
  Const OCM__BASE = WM_USER + &H1C00
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case uMsg
    Case WM_NOTIFYFORMAT
      ' give the control a chance to request Unicode notifications
      lRet = SendMessageAsLong(wParam, OCM__BASE + uMsg, wParam, lParam)

      bCallDefProc = False
  End Select

StdHandler_Ende:
  HandleMessage_Form = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in frmMain.HandleMessage_Form: ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function

Private Sub btnAktionAusführen_Click()
    Dim epm As New EasyPopupMenu                                'Gerbing 09.02.2007
        
    epm.AddMenuItem LoadResString(3005 + Sprache), MF_STRING, 1     'Öffnen der mit der Dateinamen-Erweiterung verknüpften Anwendung für die aktuelle Datei
    epm.AddMenuItem LoadResString(3119 + Sprache), MF_STRING, 2     'Öffne das Druckprogramm für die aktuelle Datei
    epm.AddMenuItem LoadResString(3120 + Sprache), MF_STRING, 3     'Email mit Anhang senden'
    epm.AddMenuItem LoadResString(3121 + Sprache), MF_STRING, 4     'Öffne ein Explorer-Fenster, wo die aktuelle Datei markiert ist
    epm.AddMenuItem LoadResString(3112 + Sprache), MF_STRING, 5     'Verschiebe alle mit Häkchen markierten Dateien in den Papierkorb    'Gerbing 23.01.2010
    epm.AddMenuItem LoadResString(3130 + Sprache), MF_STRING, 6     'Kopiere alle mit Häkchen markierten Dateien
    epm.AddMenuItem LoadResString(3154 + Sprache), MF_STRING, 7     'Zeige Geo-Position         'Gerbing 27.08.2015
    epm.AddMenuItem LoadResString(3187 + Sprache), MF_STRING, 8     'Ändern alle mit Häkchen markierten Dateinamen
    
    Call WeiterAnShellExecute(epm.TrackMenu(Me.hwnd))
    epm.DeleteMenu
    Set epm = Nothing
End Sub

Public Sub DateinamenÄndern()                                      'Gerbing 28.01.2018
    Dim strPath As String
    Dim strFile As String
    Dim strSavePath As String
    Dim DateinamenErweiterung As String
    Dim blnHäkchenGefunden As Boolean
    Dim i As Long

    blnHäkchenGefunden = False
    For i = 0 To ExLVwU.ListItems.Count - 1
        If ExLVwU.ListItems(i).StateImageIndex = 2 Or ExLVwU.ListItems(i).StateImageIndex = 4 Then
            blnHäkchenGefunden = True
            Call file_split(ListBoxForm.ExLVwU.ListItems(i), strPath, strFile, DateinamenErweiterung)
            DiashowForm.gstrFolder = strPath
            If i <> 0 Then
                If strSavePath <> strPath Then
                    'MsgBox "Wenn mehr als ein Ordner in der Liste der Dateinamen auftaucht, ist 'Dateinamen ändern' nicht möglich"
                    MsgBox LoadResString(1099 + Sprache)
                    Exit Sub
                End If
            End If
            strSavePath = strPath
       End If
    Next
    If blnHäkchenGefunden = False Then
        'MsgBox "Sie haben keine Dateien ausgewählt"            'Gerbing 13.03.2018
        MsgBox LoadResString(2303 + Sprache)
        Exit Sub
    End If
    If gblnSubdirectories = True Then
        'MsgBox "Für Dateinamen ändern darf nicht 'Mit Unterordnern' ausgewählt sein"
        MsgBox LoadResString(3188 + Sprache)
        Exit Sub
    End If
    frmNamenÄndern.Show 1
End Sub

Public Sub LöschenMarkierte()
    Dim i As Long
    Dim n As Long
    Dim blnNichtlöschbareGefunden As Boolean
    Dim blnHäkchenGefunden As Boolean
    Dim msg As String
    Dim antwort As Long
    Dim rc As Long
    Dim max As Long

    blnHäkchenGefunden = False
    For i = 0 To ExLVwU.ListItems.Count - 1
        If ExLVwU.ListItems(i).StateImageIndex = 2 Or ExLVwU.ListItems(i).StateImageIndex = 4 Then
            blnHäkchenGefunden = True
       End If
    Next
    If blnHäkchenGefunden = False Then Exit Sub
    '----------------------------------------------------------------------
    'msg = "Wollen Sie wirklich die mit Häkchen markierten Dateien löschen?"
    'msg = "Wollen Sie wirklich die mit Häkchen markierten Dateien in den Papierkorb verschieben?"  'Gerbing 23.01.2010

    msg = LoadResString(3116 + Sprache)
    antwort = MsgBox(msg, vbDefaultButton1 + vbYesNo)
    If antwort = vbNo Then Exit Sub
    '----------------------------------------------------------------------
    blnNichtlöschbareGefunden = False
    i = 0
    Do
        If ExLVwU.ListItems.Count <> 0 Then
            If ExLVwU.ListItems(i).StateImageIndex = 2 Or ExLVwU.ListItems(i).StateImageIndex = 4 Then
                'es ist ein Häkchen gesetzt
                rc = file_delete(ExLVwU.ListItems(i), True, True)               'undo, silent
                If rc <> 0 Then
                    'Löschen war erfolgreich
                    '-----------------------------------------------------------------------------------------
                    'Jetzt denselben Dateinamen aus Listbox Diashowform.List1 oder DiashowForm.List1Unsorted
                    'entfernen wie dann aus ListBoxNeu
                    If DiashowForm.chkSortAsDragAndDrop.Value = 1 Then    'Gerbing 22.01.2010
                        max = DiashowForm.List1UnsortedU.ListItems.Count - 1
                        n = 0
                        Do
                            On Error Resume Next                                                    'Gerbing 20.06.2015
                            Err.Number = 0                                                          'Gerbing 20.06.2015
                            If DiashowForm.List1UnsortedU.ListItems(n) = ExLVwU.ListItems(i) Then
                                DiashowForm.List1UnsortedU.ListItems.Remove n
                                If max = 0 Then
                                    Exit Do
                                Else
                                    max = max - 1
                                End If
                            End If
                            If n < max Then
                                n = n + 1
                            Else
                                Exit Do
                            End If
                            If Err.Number <> 0 Then                                                 'Gerbing 20.06.2015
                                Exit Do                                                             'Gerbing 20.06.2015
                            End If                                                                  'Gerbing 20.06.2015
                        Loop
                    Else
                        max = DiashowForm.List1U.ListItems.Count - 1
                        n = 0
                        Do
                            On Error Resume Next                                                    'Gerbing 20.06.2015
                            Err.Number = 0                                                          'Gerbing 20.06.2015
                            If DiashowForm.List1U.ListItems(n) = ExLVwU.ListItems(i) Then
                                DiashowForm.List1U.ListItems.Remove n
                                If max = 0 Then
                                    Exit Do
                                Else
                                    max = max - 1
                                End If
                            End If
                            If n < max Then
                                n = n + 1
                            Else
                                Exit Do
                            End If
                            If Err.Number <> 0 Then                                                 'Gerbing 20.06.2015
                                Exit Do                                                             'Gerbing 20.06.2015
                            End If                                                                  'Gerbing 20.06.2015
                        Loop
                    End If
                    On Error GoTo 0                                                                 'Gerbing 20.06.2015
                    ExLVwU.ListItems.Remove (i)
                    ExLVwUIndex = ExLVwUIndex - 1                                           'Gerbing 28.05.2013
                    '------------------------------------
                Else
                    'Löschen war nicht erfolgreich
                    blnNichtlöschbareGefunden = True
                    i = i + 1                                                               'Gerbing 12.03.2008
                End If
            Else
                'es ist kein Häkchen gesetzt
                If i < ExLVwU.ListItems.Count - 1 Then
                    i = i + 1
                Else
                    Exit Do
                End If
            End If
        Else
            Exit Do
        End If
        If i = ExLVwU.ListItems.Count Then Exit Do
    Loop
    '-------------------------------------------------
    If blnNichtlöschbareGefunden = True Then
        'MsgBox "Sie haben mindestens eine schreibgeschützte Datei markiert. Schreibgeschützte Dateien werden nicht gelöscht"
        MsgBox LoadResString(3114 + Sprache)
    End If
    If ExLVwU.ListItems.Count = 0 Then
        'MsgBox "Es gibt keine Bilder mehr anzuzeigen. Das Programm wird beendet"
        MsgBox LoadResString(3115 + Sprache)
        Set fso = Nothing
        If DiashowForm.gblnWithTitle = False Then
            Call DiashowForm.MDIFormMitTitle
        End If
        End
    End If
    Call DiashowForm.EsIstF8                                                            'Gerbing 25.08.2013
End Sub

Public Sub KopierenMarkierte()
    Dim i As Long
    Dim n As Long
    Dim blnHäkchenGefunden As Boolean
    Dim msg As String
    Dim antwort As Long
    Dim folder As String
    Dim Prompt As String
    Dim start As Long
    Dim pos1 As Long
    Dim pos2 As Long
    Dim rc As Boolean

    blnHäkchenGefunden = False
    For i = 0 To ExLVwU.ListItems.Count - 1
        If ExLVwU.ListItems(i).StateImageIndex = 2 Or ExLVwU.ListItems(i).StateImageIndex = 4 Then
             blnHäkchenGefunden = True
        End If
    Next
    If blnHäkchenGefunden = False Then Exit Sub
    '----------------------------------------------------------------------
    'Folder-Dialog
    'Prompt = "C:\"
    Prompt = ""
    folder = BrowseForFolder("Folder", Prompt, Me.hwnd, False, , True, False)                                                  'ist unicode fähig
    If folder = "" Then Exit Sub
    Me.MousePointer = vbHourglass                                                           'Gerbing 29.07.2007
    i = 0
    Do
        If ExLVwU.ListItems.Count <> 0 Then
            If ExLVwU.ListItems(i).StateImageIndex = 2 Or ExLVwU.ListItems(i).StateImageIndex = 4 Then
                start = 1                                   'der Dateiname beginnt nach dem letzten "\"
                Do
                    pos1 = InStr(start, ExLVwU.ListItems(i), "\")
                    pos2 = InStr(pos1 + 1, ExLVwU.ListItems(i), "\")
                    If pos1 <> 0 And pos2 = 0 Then Exit Do
                    start = pos2 + 1
                Loop
                If Right(folder, 1) <> "\" Then
                    folder = folder & "\"
                End If
                rc = file_copy(ExLVwU.ListItems(i), folder & Right(ExLVwU.ListItems(i), Len(ExLVwU.ListItems(i)) - pos1))
                If rc = 0 Then                                                              'rc=0=False
                    Me.MousePointer = vbNormal                                              'Gerbing 29.07.2007
                    msg = "FileCopy from " & ExLVwU.ListItems(i) & " to " & folder & Right(ExLVwU.ListItems(i), Len(ExLVwU.ListItems(i)) - pos1) & vbNewLine
                    msg = msg & "Error number=" & Err.Number & vbNewLine
                    msg = msg & "Error text=" & Err.Description & vbNewLine
                    msg = msg & "Wollen Sie den Kopiervorgang abbrechen ?"
                    'antwort = MsgBox(msg, vbDefaultButton2 + vbYesNo)
                    antwort = MessageBoxW(0, StrPtr(msg), StrPtr("GERBING Diashow"), vbDefaultButton2 + vbYesNo)
                    If antwort = vbYes Then
                        Exit Sub
                    End If
                    Me.MousePointer = vbHourglass                                           'Gerbing 29.07.2007
                End If
                On Error GoTo 0
            End If
            If i < ExLVwU.ListItems.Count - 1 Then
                    i = i + 1
                Else
                    Exit Do
            End If
        Else
            Exit Do
        End If
        If i = ExLVwU.ListItems.Count Then Exit Do
    Loop
    ListBoxForm.Hide
    Me.MousePointer = vbNormal                                                              'Gerbing 29.07.2007
    Call DiashowForm.BildAnzeigen
End Sub

Private Sub chkExifAnzeigen_Click()                             'Gerbing 20.02.2007
    Call EXIFLesen
End Sub

Private Sub chkiptcAnzeigen_Click()                             'Gerbing 20.02.2007
    Dim tempDateiname As String
    Dim rc As Boolean
    
    If StrComp(Right(gblStrAktuellGezeigtesBild, 3), "JPG", vbTextCompare) <> 0 Then
        chkIptcAnzeigen.Value = 0
        Exit Sub                                                'Gerbing 27.05.2015
    End If
    
    'chkExifAnzeigen.Value = 0                                  'Gerbing 18.08.2008
    If chkIptcAnzeigen.Value = 0 Then
        txtIPTCInfo.Visible = False
        txtEXIFInfo.Visible = False
    Else
        chkExifAnzeigen.Value = 0                                  'Gerbing 18.08.2008
        txtIPTCInfo.Visible = True
        txtEXIFInfo.Visible = False                             'Gerbing 07.05.2007
        If Not ListBoxForm.ExLVwU.ListItems(ExLVwUIndex) Is Nothing Then
            If gblStrAktuellGezeigtesBild = ListBoxForm.ExLVwU.ListItems(ExLVwUIndex) Then
                'Call BastleIPTCUni                              'Gerbing 04.03.2013
                IPTCItemsDelimiter = ";"
                rc = LeseIPTC(gblStrAktuellGezeigtesBild, txtIPTCInfo, True)  'mit Ausgabe in txtIPTCInfo
            Else
                Call StimmtNichtÜberein
            End If
        End If
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 0 Then Exit Sub                            'Gerbing 25.08.2013
    'Me.Hide                                                'Gerbing 11.10.2016 auskommentiert 'sonst flackert es beim Drücken jeder Taste
    Call DiashowForm.Form_KeyDown(KeyCode, Shift)           'Gerbing 19.12.2012
End Sub

Private Sub Form_Load()
    Dim Faktor As Integer
    Dim Pixel As Integer
    Dim rc As Long
    Dim i As Long
    
    ' this is required to make the Timosoft control work as expected in unicode
    'Subclass

    Call AnpassenNutzerWunsch(Me)                               'Gerbing 11.03.2017
    
    'ListBoxForm.BorderStyle = vbSizableToolWindow    muß im Design-mode gesetzt werden
    btnAktionAusführen.Caption = LoadResString(3117 + Sprache)  '&Aktion wählen und ausführen...
    'Horizontale Scrollbar an txtIPTCInfo anbringen        'Gerbing 09.11.2006
    Faktor = 4
    Pixel = txtIPTCInfo.Width / Screen.TwipsPerPixelX * Faktor
    rc = SendMessage(txtIPTCInfo.hwnd, LB_SETHORIZONTALEXTENT, Pixel, 0)
    
    chkExifAnzeigen.Caption = LoadResString(1116 + Sprache) 'Alle Informationen
    chkIptcAnzeigen.Caption = LoadResString(1117 + Sprache) 'IPTC-Felder
    btnAlleHäkchenSetzen.ToolTipText = LoadResString(2559 + Sprache) 'alle markieren
    btnAlleHäkchenEntfernen.ToolTipText = LoadResString(2560 + Sprache) 'alle Markierungen entfernen
    ListBoxForm.Width = MDIForm1.Width - 500                    'Gerbing 03.03.2013
    
    'ExLVwU.Columns.Add "Column 1", , 500, 100, alLeft
    'ExLVwU.Columns.Add "Column 1", ExLVwU.Width - 10, , , 1
    ExLVwU.Columns.Add "Column 1"
    formCaption Me.hwnd, Me.Caption      'Gerbing 10.06.2013
End Sub

Private Sub Form_Resize()
   On Error Resume Next
    ExLVwU.Width = ListBoxForm.Width - 400 - FrameExifIptc.Width
    FrameExifIptc.Left = ListBoxForm.Width - FrameExifIptc.Width
    ExLVwU.Height = ListBoxForm.Height - 1200   '- 865
    FrameExifIptc.Height = ExLVwU.Height + 565
    txtIPTCInfo.Height = ExLVwU.Height - 250
    txtEXIFInfo.Height = ExLVwU.Height - 250
End Sub

Private Sub Form_Terminate()
  If hImgLst Then ImageList_Destroy hImgLst
  If hStateImgLst Then ImageList_Destroy hStateImgLst
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    Dim msg As String
    Dim antwort As Long
    
    If Not UnSubclassWindow(Me.hwnd, EnumSubclassID.escidFrmMain) Then
        Debug.Print "UnSubclassing failed!"
    End If
    
    '---------------------------------------------------------------'Gerbing 07.12.2009
    If gblnHäkchenGefunden = False Then                             'Gerbing 19.10.2017
        For i = 0 To ListBoxForm.ExLVwU.ListItems.Count - 1
                If ExLVwU.ListItems(i).StateImageIndex = 2 Or ExLVwU.ListItems(i).StateImageIndex = 4 Then
                    gblnHäkchenGefunden = True
                End If
            Next
        If gblnHäkchenGefunden = True Then
            'msg = "Es gibt mit Häkchen markierte Dateien, wollen Sie diese wirklich ignorieren?"
            msg = LoadResString(3149 + Sprache)
            antwort = MsgBox(msg, vbDefaultButton1 + vbYesNo)
            If antwort = vbYes Then
                Exit Sub
            Else
                Cancel = True
            End If
        End If
    End If
End Sub

Public Sub WeiterAnShellExecute(ID)
    Dim retval As Long
   
    If ID = 0 Then Exit Sub     'Bei ID = 0 wurde keine Aktion ausgewählt           'Gerbing 16.02.2007
    ID = ID - 1

'   Werte für Combo1.Listindex
'   0 = 'Öffnen der mit der Dateinamen-Erweiterung verknüpften Anwendung für die aktuelle Datei
'   1 = Öffne das Druckprogramm für die aktuelle Datei
'   2 = Email mit Anhang senden
'   3 = Öffne ein Explorer-Fenster, wo die aktuelle Datei markiert ist
'   4 = Lösche alle mit Häkchen markierten Dateien in den Papierkorb                'Gerbing 23.01.2010
'   5 = Kopiere alle mit Häkchen markierten Dateien
'   6 = Zeige Geo-Position                                                          'Gerbing 27.08.2015
'   7 = Ändern alle mit Häkchen markierten Dateinamen                               'Gerbing 28.01.2018
    Select Case ID
    'Select Case Combo1.ListIndex
        Case 0
            '0 = Öffnen der mit 'jpg' verknüpften Anwendung für die aktuelle Datei
            Call ÖffneVerknüpfteAnwendung
        Case 1
            '1 = Öffne das Druckprogramm für die aktuelle Datei
'            retVal = RunShellExecute(Me.hWnd, "Print", ExLVwU.ListItems(ExLVwUIndex), vbNullString, vbNullString, 1)
'            retVal = RunShellExecute(Me.hWnd, "Print", ExLVwU.ListItems(ExLVwUIndex), vbNullString, CurDir$, 1)
'            If retVal <= 32 Then
'                MsgBox LoadResString(3122 + Sprache) 'Es wurde kein geeignetes Druckprogramm gefunden um diese Datei auszudrucken
'            End If
            Shell ("rundll32.exe SHELL32,OpenAs_RunDLL " & ExLVwU.ListItems(ExLVwUIndex))                                'Gerbing 07.08.2013
        Case 2
            '2 = Öffne das Fenster 'Neue Email senden'
            'retVal = RunShellExecute(Me.hWnd, "open", "mailto:xxx@yyy.zzz?", vbNull, vbNull, 1)
            Call EmailMitAnhangSenden                                                                        'Gerbing 13.08.2017
        Case 3
            '3 = Öffne ein Explorer-Fenster, wo die aktuelle Datei markiert ist
            retval = RunShellExecute(Me.hwnd, "open", "explorer.exe", "/e,/select," & ExLVwU.ListItems(ExLVwUIndex), vbNull, 1)
        Case 4                                      'Gerbing 28.08.2006
            '4 = Lösche alle mit Häkchen markierten Dateien in den Papierkorb                       'Gerbing 23.01.2010
            Call LöschenMarkierte
        Case 5
            '5 = Kopiere alle mit Häkchen markierten Dateien
            Call KopierenMarkierte
        Case 6
            '6 = Zeige GEO-Position                                                                 'Gerbing 27.08.2015
            Call ZeigeGeoPosition                                                                   'Gerbing 29.01.2018
        Case 7
        'Ändern alle mit Häkchen markierten Dateinamen                              'Gerbing 28.01.2018
        Call DateinamenÄndern                                                       'Gerbing 28.01.2018
    End Select
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

Private Sub StimmtNichtÜberein()
'    txtEXIFInfo = "Der Dateiname in der Dateinamenliste" & vbNewLine
'    txtEXIFInfo = txtEXIFInfo & "und der Name der aktuell angezeigten" & vbNewLine
'    txtEXIFInfo = txtEXIFInfo & "Datei sind unterschiedlich."                      'Gerbing 29.07.2007
'    txtEXIFInfo = txtEXIFInfo & "Sie müssen Doppelklicken oder"                    'Gerbing 16.08.2007
'    txtEXIFInfo = txtEXIFInfo & "die Enter-Taste drücken, um die"
'    txtEXIFInfo = txtEXIFInfo & "ausgewählte Datei anzuzeigen."
    '------------------------------------------------------------------------------
    txtEXIFInfo = LoadResString(1225 + Sprache) & vbNewLine
    txtEXIFInfo = txtEXIFInfo & LoadResString(1226 + Sprache) & vbNewLine
    txtEXIFInfo = txtEXIFInfo & LoadResString(1227 + Sprache) & vbNewLine
    txtEXIFInfo = txtEXIFInfo & LoadResString(1228 + Sprache) & vbNewLine
    txtEXIFInfo = txtEXIFInfo & LoadResString(1229 + Sprache) & vbNewLine
    txtEXIFInfo = txtEXIFInfo & LoadResString(1230 + Sprache)
    txtEXIFInfo.Visible = True
End Sub

Private Function GEOKoordinatenUmrechnen()
    'rc = 0 ohne Fehler
    'rc = 1 Fehler
    Dim pos1 As Long    'erster Bindestrich
    Dim pos2 As Long    'zweiter Bindestrich
    Dim pos3 As Long    'komma
    Dim Grad As String
    Dim Minuten As String
    Dim Sekunden As String
    Dim Hundertstel As String
    Dim lngSekunden
    Dim lngDezimalanteil
    Dim strLatitude As String
    Dim strLongitude As String
    
    GEOKoordinatenUmrechnen = 0
    
    'gstrGEOPosition zusammensetzen                                                 'Gerbing 02.09.2016
    'zB 50.83266,12.45735
    gstrGEOPosition = ""
    gstrLat = ""                                                                    'Gerbing 29.09.2018
    gstrLong = ""                                                                   'Gerbing 29.09.2018
    If GPSLatitudeRef <> "N" Then
        gstrGEOPosition = "-"
    End If
    GPSLatitude = Replace(GPSLatitude, ",", ".")
    gstrLat = gstrGEOPosition & GPSLatitude                                         'Gerbing 29.09.2018
    gstrGEOPosition = gstrGEOPosition & GPSLatitude & ","
    If GPSLongitudeRef <> "E" Then
        gstrGEOPosition = gstrGEOPosition & "-"
        gstrLong = "-"                                                              'Gerbing 29.09.2018
    End If
    GPSLongitude = Replace(GPSLongitude, ",", ".")
    gstrGEOPosition = gstrGEOPosition & GPSLongitude
    gstrLong = gstrLong & GPSLongitude                                              'Gerbing 29.09.2018
    Exit Function
End Function

Public Sub EmailMitAnhangSenden()                                          'Gerbing 13.08.2017
    ' Outlook Applikation
    Dim ool As Outlook.Application
    Dim oInspector As Outlook.Inspector
    Dim oMail As Outlook.MailItem
    Dim myattachments As Variant
    
    ' Für Inputbox "EMailadresse-Änderung"
    Dim Mldg, Titel, Voreinstellung, MailAdress
    
    
    '     ' Adresse anzeigen und Änderung ermöglichen
    '     Mldg = "Ist die angegebene Emailadresse richtig?"
    '     Titel = "Mailadresse"
    '     Voreinstellung = "trallala@gmx.de"
    '     MailAdress = InputBox(Mldg, Titel, Voreinstellung)
    MailAdress = LoadResString(1239 + Sprache)                              '"Bitte eingeben"
    
    ' Wurde Abbrechen gedrückt, dann alles beenden
    If MailAdress = "" Then Exit Sub
    
    ' Verweis zu Outlook + neue Nachricht
    On Error GoTo Outlookfehlt
    Set ool = CreateObject("Outlook.Application")
    Set oMail = ool.CreateItem(olMailItem)
    Set myattachments = oMail.Attachments
    
    ' Befreff-Zeile
    oMail.Subject = LoadResString(1239 + Sprache)                           '"Bitte eingeben"
    
    ' An-Zeile (Empfänger)
    oMail.To = MailAdress
    'oMail.Recipients.ResolveAll                      'hier kommt error 427
    oMail.Display
    
    ' Texteingabe (Nachricht selbst)
    oMail.Body = LoadResString(1239 + Sprache)                              '"Bitte eingeben"
    
    
    ' Anhang
    ' Nachfolgend ein Beispiel. Suchen Sie sich eine Datei auf
    ' Ihrem Rechner aus - vollständiger Pfad muß mitangegeben
    ' sein.
    ' Es können auch weitere Dateien angegeben werden.
    ' Hierzu einfach mit myattachments.Add "???" fortsetzen.
    myattachments.Add "" & ExLVwU.ListItems(ExLVwUIndex) & ""
    
    ' Speicher freigeben
    Set ool = Nothing
    Set oInspector = Nothing
    Set oMail = Nothing
    Exit Sub
Outlookfehlt:
    MsgBox LoadResString(1240 + Sprache)                                    '"Outlook ist nicht installiert"
End Sub

Public Sub ZeigeGeoPosition()                                                              'Gerbing 29.01.2018
    Dim pos As Long
    Dim pos1 As Long
    Dim pos2 As Long
    Dim rc As Long
    Dim url As String
    Const SW_SHOWNORMAL = 1                                                             'Gerbing 16.10.2018
    Const SW_SHOWMAXIMIZED = 3                                                          'Gerbing 16.10.2018
    
    'aber nur für JPG files
'            If StrComp(Right(gblStrAktuellGezeigtesBild, 3), "JPG", vbTextCompare) <> 0 Then
'                Exit Sub
'            End If
    DiashowForm.EXF.ImageFile = ListBoxForm.ExLVwU.ListItems(ListBoxForm.ExLVwUIndex) 'set the image file property
    If gstrLatXMP <> "" And gstrLongXMP <> "" Then                                              'Gerbing 08.04.2019
        'es gibt GEO Positionen im XMP-Abschnitt
        rc = GEOKoordinatenUmrechnenXMP
    Else
        'suche die GEO Positinen im EXIF-Abschnitt
        ListBoxForm.txtEXIFInfo.Visible = True
        If ListBoxForm.ExLVwU.ListItems(ListBoxForm.ExLVwUIndex) = "" Then
            ListBoxForm.txtEXIFInfo.Text = LoadResString(1483 + Sprache)  'keine Datei markiert
        Else
            '
            'EXF.ListInfo ist ein String mit vbCrLf
            '
            ListBoxForm.txtEXIFInfo.Text = DiashowForm.EXF.ListInfo 'list all tags into the text box
        End If
        pos = InStr(1, ListBoxForm.txtEXIFInfo.Text, "GPSLatitudeRef:", vbTextCompare)
        If pos = 0 Then
            MsgBox LoadResString(3155 + Sprache)  'keine GEO-Positionen vorhanden
            Exit Sub
        End If
        GPSLatitudeRef = Mid(ListBoxForm.txtEXIFInfo.Text, pos + 16, 1)
        pos = InStr(1, ListBoxForm.txtEXIFInfo.Text, "GPSLatitude:", vbTextCompare)
        If pos = 0 Then                                                                         'Gerbing 07.03.2016
            MsgBox LoadResString(3155 + Sprache)  'keine GEO-Positionen vorhanden
            Exit Sub
        End If
        pos1 = InStr(pos, ListBoxForm.txtEXIFInfo.Text, ":", vbTextCompare)         'suche den Doppelpunkt Gerbing 02.09.2016
        pos2 = InStr(pos1, ListBoxForm.txtEXIFInfo.Text, vbNewLine)                 'suche das Zeilenende
        GPSLatitude = Mid(ListBoxForm.txtEXIFInfo.Text, pos1 + 2, pos2 - pos1 - 2)
        pos = InStr(1, ListBoxForm.txtEXIFInfo.Text, "GPSLongitudeRef:", vbTextCompare)
        GPSLongitudeRef = Mid(ListBoxForm.txtEXIFInfo.Text, pos + 17, 1)
        pos = InStr(1, ListBoxForm.txtEXIFInfo.Text, "GPSLongitude:", vbTextCompare)
        pos1 = InStr(pos, ListBoxForm.txtEXIFInfo.Text, ":", vbTextCompare)         'suche den Doppelpunkt
        pos2 = InStr(pos1, ListBoxForm.txtEXIFInfo.Text, vbNewLine)                 'suche das Zeilenende
        GPSLongitude = Mid(ListBoxForm.txtEXIFInfo.Text, pos1 + 2, pos2 - pos1 - 2)
        rc = GEOKoordinatenUmrechnen
    End If
    If rc = 0 Then
        'frmGEOPosition.Show 1
        frmStrgG.Show 1                                                             'Gerbing 16.10.2018
        Select Case glngStrgG                                                       'Gerbing 15.04.2020
            Case 1
                frmMap.Show 1
            Case 2
                url = "http://www.openstreetmap.org/?mlat=" & gstrLat & "&mlon=" & gstrLong & "&zoom=16&layers=M?force=tt&hl=de-AT" 'Gerbing 16.10.2018
                ' "Execute" the URL to make the default browser display it.         'Gerbing 16.10.2018
                ShellExecute ByVal 0&, "open", url, _
                    vbNullString, vbNullString, SW_SHOWNORMAL                       'Gerbing 16.10.2018
            Case 3                                                                  'Gerbing 15.04.2020
                'url = "https://maps.google.com/maps?q=50.8359%2C12.9229"
                url = "https://maps.google.com/maps?q=" & gstrLat & "%2C" & gstrLong
                ' "Execute" the URL to make the default browser display it.
                ShellExecute ByVal 0&, "open", url, _
                    vbNullString, vbNullString, SW_SHOWNORMAL
        End Select                                                                  'Gerbing 15.04.2020
    End If
End Sub

Public Function GEOKoordinatenUmrechnenXMP()                                       'Gerbing 08.04.2019
    'zB gstrLatXMP 50,38.7309456N -> 50.64551575
    'zB gstrLongXMP 11,53.9826786E -> 11.89971130
    'Das ist nötig damit die GEO-Positionen von OpenStreetMap verstanden werden
    Dim Grad As String
    Dim Minuten As String                                                   'Gerbing 04.07.2019
    Dim MinutenDouble As Double                                             'Gerbing 04.07.2019
    Dim ESWN As String                                                      'East South West Nord
    Dim Ergebnis As String
    Dim pos As Integer
    Dim locale_id As Long                                                   'Gerbing 04.07.2019
    
    GEOKoordinatenUmrechnenXMP = 0
    pos = InStr(1, gstrLatXMP, ",")                                         'das "," kommt in deutscher und englischer Systemsprache
    Grad = Mid(gstrLatXMP, 1, pos - 1)
    Minuten = Mid(gstrLatXMP, pos + 1, Len(gstrLatXMP) - pos - 1)
    'Wenn Komma als Dezimaltrennzeichen verwendet wird, muss der Punkt im String Minuten in Komma verwandelt werden
    'sonst kommt bei MinutenDouble / 60 Ergebnis=0
    If LocaleInfo(locale_id, LOCALE_SDECIMAL) = "," Then
        Minuten = Replace(Minuten, ".", ",")
    End If
    ESWN = Mid(gstrLatXMP, Len(gstrLatXMP), 1)
    MinutenDouble = CDbl(Minuten)
    MinutenDouble = MinutenDouble / 60
    Ergebnis = Grad + MinutenDouble
    Ergebnis = Replace(Ergebnis, ",", ".")                                  'Gerbing 04.07.2019
    gstrLat = ""
    If ESWN <> "N" Then
        gstrLat = "-"                                                       '- auf der Südhalbkugel
    End If
    gstrLat = gstrLat & Ergebnis                                            'Gerbing 04.07.2019
    '---------------------------
    pos = InStr(1, gstrLongXMP, ",")
    Grad = Mid(gstrLongXMP, 1, pos - 1)
    Minuten = Mid(gstrLongXMP, pos + 1, Len(gstrLongXMP) - pos - 1)
    'Wenn Komma als Dezimaltrennzeichen verwendet wird, muss der Punkt im String Minuten in Komma verwandelt werden
    'sonst kommt bei MinutenDouble / 60 Ergebnis=0
    If LocaleInfo(locale_id, LOCALE_SDECIMAL) = "," Then
        Minuten = Replace(Minuten, ".", ",")
    End If
    ESWN = Mid(gstrLongXMP, Len(gstrLongXMP), 1)
    MinutenDouble = CDbl(Minuten)
    MinutenDouble = MinutenDouble / 60
    Ergebnis = Grad + MinutenDouble
    Ergebnis = Replace(Ergebnis, ",", ".")                                  'Gerbing 04.07.2019
    gstrLong = ""
    If ESWN <> "E" Then
        gstrLong = "-"                                                       '- westlich von Greenwich
    End If
    gstrLong = gstrLong & Ergebnis
End Function

Public Sub ÖffneVerknüpfteAnwendung()                                          'Gerbing 29.01.2018
    Dim retval As Long
    Dim DateinamenErweiterung As String
    Dim intLänge As Integer
    Dim ErrorText As String
    Dim msg As String

    retval = RunShellExecute(Me.hwnd, "open", ExLVwU.ListItems(ExLVwUIndex), vbNull, vbNull, 1)
    If retval <= 32 Then
        If Mid(ExLVwU.ListItems(ExLVwUIndex), Len(ExLVwU.ListItems(ExLVwUIndex)) - 3, 1) = "." Then
            intLänge = 3
        End If
        If Mid(ExLVwU.ListItems(ExLVwUIndex), Len(ExLVwU.ListItems(ExLVwUIndex)) - 4, 1) = "." Then
            intLänge = 4
        End If
        If Mid(ExLVwU.ListItems(ExLVwUIndex), Len(ExLVwU.ListItems(ExLVwUIndex)) - 5, 1) = "." Then
            intLänge = 5
        End If
        DateinamenErweiterung = Right(ExLVwU.ListItems(ExLVwUIndex), intLänge)
        ErrorText = GetShellError(retval)           'Gerbing 20.08.2008
        msg = "Errortext=" & ErrorText & vbNewLine
        msg = msg & "Errornr=" & retval & vbNewLine & vbNewLine
        
        msg = msg & ExLVwU.ListItems(ExLVwUIndex) & vbNewLine
        'Msg = Msg & "Diese Datei kann nicht geöffnet werden." & vbNewLine & vbNewLine
        msg = msg & LoadResString(1376 + Sprache) & vbNewLine & vbNewLine
        
        'Msg = Msg & "Entweder die Datei existiert nicht," & vbNewLine & vbNewLine
        msg = msg & LoadResString(2208 + Sprache) & vbNewLine & vbNewLine
        
        'Msg = Msg & "oder es ist keine Anwendung mit der" & vbNewLine
        msg = msg & LoadResString(1378 + Sprache) & vbNewLine
        'Msg = Msg & "Dateinamen-Erweiterung(Datei-Typ) " & DateinamenErweiterung & " verknüpft." & vbNewLine
        msg = msg & LoadResString(1379 + Sprache) & DateinamenErweiterung & LoadResString(1380 + Sprache) & vbNewLine
        'Msg = Msg & "Wählen Sie selbst eine geignete Anwendung, zB mittels Windows-Explorer" & vbNewLine
        msg = msg & LoadResString(2012 + Sprache) & vbNewLine
        'Msg = Msg & "Rechtklicken auf den Dateiname -> Öffnen mit... -> Programm auswählen"
        'msg = msg & LoadResString(2013 + Sprache)
        MessageBoxW 0, StrPtr(msg), StrPtr("GERBING Diashow"), vbInformation
        'MsgBox msg
    End If
End Sub

' Return a piece of locale information.
Private Function LocaleInfo(ByVal locale As Long, ByVal lc_type As Long) As String
Dim Length As Long
Dim buf As String * 1024

    Length = GetLocaleInfo(locale, lc_type, buf, Len(buf))
    LocaleInfo = Left$(buf, Length - 1)
End Function

