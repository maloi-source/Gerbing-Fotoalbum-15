VERSION 5.00
Object = "{A8F9B8E7-E699-4FCE-A647-72C877F8E632}#1.8#0"; "editctlsu.ocx"
Begin VB.Form frmNamenÄndern 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Dateinamen ändern"
   ClientHeight    =   3000
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   11364
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   11364
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton btnAbbrechen 
      Caption         =   "Abbrechen"
      Height          =   492
      Left            =   5520
      TabIndex        =   6
      Top             =   2280
      Width           =   1692
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   492
      Left            =   3480
      TabIndex        =   5
      Top             =   2280
      Width           =   1692
   End
   Begin EditCtlsLibUCtl.TextBox txtGewählterOrdner 
      Height          =   492
      Left            =   3480
      TabIndex        =   4
      Top             =   120
      Width           =   7692
      _cx             =   13568
      _cy             =   868
      AcceptNumbersOnly=   0   'False
      AcceptTabKey    =   0   'False
      AllowDragDrop   =   -1  'True
      AlwaysShowSelection=   0   'False
      Appearance      =   1
      AutoScrolling   =   2
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
      MultiLine       =   0   'False
      OLEDragImageStyle=   0
      PasswordChar    =   0
      ProcessContextMenuKeys=   -1  'True
      ReadOnly        =   -1  'True
      RegisterForOLEDragDrop=   0   'False
      RightMargin     =   -1
      RightToLeft     =   0
      ScrollBars      =   0
      SelectedTextMousePointer=   0
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      UseCustomFormattingRectangle=   0   'False
      UsePasswordChar =   0   'False
      UseSystemFont   =   0   'False
      CueBanner       =   "frmNamenÄndern.frx":0000
      Text            =   "frmNamenÄndern.frx":0020
   End
   Begin EditCtlsLibUCtl.TextBox txtNeuerDateiname 
      Height          =   492
      Left            =   3480
      TabIndex        =   3
      Top             =   1560
      Width           =   7692
      _cx             =   13568
      _cy             =   868
      AcceptNumbersOnly=   0   'False
      AcceptTabKey    =   0   'False
      AllowDragDrop   =   -1  'True
      AlwaysShowSelection=   0   'False
      Appearance      =   1
      AutoScrolling   =   2
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
      MultiLine       =   0   'False
      OLEDragImageStyle=   0
      PasswordChar    =   0
      ProcessContextMenuKeys=   -1  'True
      ReadOnly        =   0   'False
      RegisterForOLEDragDrop=   0   'False
      RightMargin     =   -1
      RightToLeft     =   0
      ScrollBars      =   0
      SelectedTextMousePointer=   0
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      UseCustomFormattingRectangle=   0   'False
      UsePasswordChar =   0   'False
      UseSystemFont   =   0   'False
      CueBanner       =   "frmNamenÄndern.frx":0040
      Text            =   "frmNamenÄndern.frx":0060
   End
   Begin VB.Label lblNeuerDateiname 
      Caption         =   "Neuer Dateiname:"
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1812
   End
   Begin VB.Label lblGewählterOrdner 
      Caption         =   "gewählter Ordner:"
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3132
   End
   Begin VB.Label lblHinweisNummer 
      Caption         =   "Alle Dateien im gewählten Ordner bekommen den neuen Dateinamen gefolgt von einer aufsteigenden Nummer"
      Height          =   372
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Neuer Dateiname_001.jpg"
      Top             =   960
      Width           =   11052
   End
End
Attribute VB_Name = "frmNamenÄndern"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAbbrechen_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    Dim strPath As String
    Dim strFile As String
    Dim DateinamenErweiterung As String
    Dim NameNeu As String
    Dim LaufendeNr As Long
    Dim VorNullen As String
    Dim EinfügNullen As String
    Dim i As Long
    Dim rc As Long

    If txtNeuerDateiname.Text = "" Then Exit Sub
    VorNullen = ""
    If ListBoxForm.ExLVwU.ListItems.Count > 10 Then VorNullen = "0"
    If ListBoxForm.ExLVwU.ListItems.Count > 100 Then VorNullen = "00"
    If ListBoxForm.ExLVwU.ListItems.Count > 1000 Then VorNullen = "000"
    If ListBoxForm.ExLVwU.ListItems.Count > 10000 Then VorNullen = "0000"
    If ListBoxForm.ExLVwU.ListItems.Count > 100000 Then VorNullen = "00000"
    For i = 0 To ListBoxForm.ExLVwU.ListItems.Count - 1
        If ListBoxForm.ExLVwU.ListItems(i).StateImageIndex = 2 Or ListBoxForm.ExLVwU.ListItems(i).StateImageIndex = 4 Then
            'file_split splits a complete file name into directory, file name and extension:
            Call file_split(ListBoxForm.ExLVwU.ListItems(i), strPath, strFile, DateinamenErweiterung)
            LaufendeNr = LaufendeNr + 1
            If LaufendeNr < 10 Then
                EinfügNullen = Mid(VorNullen, 1, Len(VorNullen))
            ElseIf LaufendeNr < 100 Then
                EinfügNullen = Mid(VorNullen, 1, Len(VorNullen) - 1)
            ElseIf LaufendeNr < 1000 Then
                EinfügNullen = Mid(VorNullen, 1, Len(VorNullen) - 2)
            ElseIf LaufendeNr < 10000 Then
                EinfügNullen = Mid(VorNullen, 1, Len(VorNullen) - 3)
            ElseIf LaufendeNr < 100000 Then
                EinfügNullen = Mid(VorNullen, 1, Len(VorNullen) - 4)
            End If
            NameNeu = txtGewählterOrdner.Text & txtNeuerDateiname.Text & "_" & EinfügNullen & CStr(LaufendeNr) & "." & DateinamenErweiterung
            rc = NameAs(ListBoxForm.ExLVwU.ListItems(i), NameNeu)  'NameAlt,NameNeu
            If rc = 0 Then                                          'Gerbing 13.03.2018
                Unload Me                                           'Gerbing 13.03.2018
                Exit Sub                                            'Gerbing 13.03.2018
            End If
       End If
    Next
    'MsgBox "Fertig"
    MsgBox LoadResString(1007 + Sprache)
    Unload Me
End Sub

Private Sub Form_Load()
    Call AnpassenNutzerWunsch(Me)
    
    Me.Caption = LoadResString(1241 + Sprache)                  'Dateinamen ändern
    lblGewählterOrdner.Caption = LoadResString(1242 + Sprache)  'Gewählter Ordner:
    lblHinweisNummer.Caption = LoadResString(1243 + Sprache)    'Alle Dateien im gewählten Ordner bekommen den neuen Dateinamen gefolgt von einer aufsteigenden Nummer
    lblNeuerDateiname.Caption = LoadResString(1244 + Sprache)   'Neuer Dateiname:
    btnAbbrechen.Caption = LoadResString(3013 + Sprache)        'Abbrechen
    lblHinweisNummer.ToolTipText = LoadResString(1245 + Sprache) 'Neuer Dateiname_001.jpg
    txtGewählterOrdner.Text = DiashowForm.gstrFolder
End Sub

 
