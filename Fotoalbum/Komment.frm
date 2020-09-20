VERSION 5.00
Object = "{A8F9B8E7-E699-4FCE-A647-72C877F8E632}#1.8#0"; "editctlsu.ocx"
Begin VB.Form KommentarForm 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   5  'Änderbares Werkzeugfenster
   Caption         =   "Fotoalbum-Kommentar"
   ClientHeight    =   3804
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   15144
   Icon            =   "Komment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3804
   ScaleWidth      =   15144
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnKommentarSpeichern 
      Caption         =   "&Kommentar speichern"
      Height          =   372
      Left            =   5160
      TabIndex        =   0
      Top             =   120
      Width           =   3492
   End
   Begin EditCtlsLibUCtl.TextBox txtKommentar 
      Height          =   3132
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   14892
      _cx             =   26268
      _cy             =   5524
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
      CueBanner       =   "Komment.frx":038A
      Text            =   "Komment.frx":03AA
   End
End
Attribute VB_Name = "KommentarForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Dim RetteInhalt As String
    Dim Msg As String
    Dim rc As Long

Private Sub btnKommentarSpeichern_Click()                                               'Gerbing 19.07.2005
    Dim strTemp As String
    Dim maximum As String
    Dim pos1 As Long

    'Kommentar nur speichern wenn sich etwas geändert hat                               'Gerbing 10.04.2016
    If frmGridAndThumb.Adodc1.Recordset(LoadResString(1030 + Sprache)) = KommentarForm.txtKommentar.Text Then
        Exit Sub
    End If
    If gstrCommandLine = "/WRITE" Then
        If Form1.KommentarFensterEinblenden = True Then
            '                                                                           'Gerbing 16.10.2014
            'Nur bei meiner privaten Datenbank gibt es exifdatetimeoriginal und nur hier soll die Gültigkeitsprüfung stattfinden
            'If Gefundenexifdatetimeoriginal = True And gblnVollversion = True Then
            If Gefundenexifdatetimeoriginal = True And gblnVollversion = True And Sprache = 0 And gblnSQLServerVersion = False Then
                strTemp = oCat.Tables("fotos").Columns("Kommentar").Properties(8)           '8=ValidationRule=Len([Kommentar])<2001
                If strTemp <> "" Then       'gibt es eine Gültigkeitsregel?
                    'es gibt eine Gültigkeitsregel
                    pos1 = InStr(1, strTemp, "<", vbTextCompare)
                    If pos1 <> 0 Then
                        maximum = Mid(strTemp, pos1 + 1, Len(strTemp) - pos1)
                        If IsNumeric(maximum) Then
                            Trim (KommentarForm.txtKommentar.Text)
                            If Not Len(KommentarForm.txtKommentar.Text) < maximum Then
                                MsgBox "Gültigkeitsregel: " & oCat.Tables("fotos").Columns("Kommentar").Properties(7)   '7=ValidationText
                                KommentarForm.txtKommentar.Text = Left(KommentarForm.txtKommentar.Text, 2000)           'begrenzen auf maximal erlaubte Bytes
                                'Zum Speichern muss der Nutzer nochmals auf 'Kommentar speichern' klicken
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
            '                                                                           'Gerbing 16.10.2014
            'frmGridAndThumb.Adodc1.Recordset.Edit
            frmGridAndThumb.Adodc1.Recordset(LoadResString(1030 + Sprache)) = KommentarForm.txtKommentar.Text
            If StrComp(Right(frmGridAndThumb.Adodc1.Recordset(LoadResString(1028 + Sprache)), 3), "jpg", vbTextCompare) = 0 Then 'Gerbing 10.04.2016
                'nur wenn die rechten 3 Bytes des Dateinamens = jpg                     'Gerbing 10.04.2016
                frmGridAndThumb.Adodc1.Recordset("IPTCPresent") = 0                     'Gerbing 10.04.2016
            End If                                                                      'Gerbing 10.04.2016
            On Error Resume Next
            frmGridAndThumb.Adodc1.Recordset.Update
            If Err.Number <> 0 Then                                                     'Gerbing 30.05.2012
                Msg = Err.Number & vbNewLine
                Msg = Msg & Err.Description
                'MsgBox Msg
                MessageBoxW 0, StrPtr(Msg), StrPtr(LoadResString(1119 + Sprache)), vbInformation
                On Error GoTo 0
            End If
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)                          'Gerbing 11.09.2014
    Select Case KeyCode
        Case vbKeyF1, vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF11
        Unload Me
        'Tastatur-Eingabe weiterreichen
        '-> Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
         Call Form1.Form_KeyDown(KeyCode, Shift)
    End Select
End Sub

Private Sub Form_Load()
    Dim iX As Long
    Dim RetVal As Long
    Dim strKommentar As String

    Call AnpassenNutzerWunsch(Me)                       'Gerbing 11.03.2017
    Me.Caption = LoadResString(1082 + Sprache)          'Fotoalbum-Kommentar Gerbing 08.11.2005
    'me.MousePointer = vbHourglass
    If glngKommentarWidth = 0 Then
        SetWindowPos hWnd, conHwndTopmost, 0, 0, 1024, 150, conSwpNoActivate Or conSwpShowWindow
    Else
        SetWindowPos hWnd, conHwndTopmost, glngKommentarLeft \ Screen.TwipsPerPixelX, glngKommentarTop \ Screen.TwipsPerPixelY, _
                        glngKommentarWidth \ Screen.TwipsPerPixelX, _
                        glngKommentarHeight \ Screen.TwipsPerPixelY, conSwpNoActivate Or conSwpShowWindow           'Gerbing 30.05.2013
    End If
    Me.Visible = False
    KommentarForm.KeyPreview = True
    btnKommentarSpeichern.Caption = LoadResString(3021 + Sprache) '&Kommentar speichern
    If gblnComeFromThumbs = False Then                                                              'Gerbing 04.05.2015
        If Not IsNull(frmGridAndThumb.rsDataGrid(LoadResString(1030 + Sprache))) Then
            strKommentar = frmGridAndThumb.rsDataGrid(LoadResString(1030 + Sprache))
            If strKommentar <> "" Then
                txtKommentar.Text = strKommentar
            End If
        End If
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
'    txtKommentar.Width = Me.Width - 200
'    txtKommentar.Height = Me.Height - 1000
    
    txtKommentar.width = Me.width - 400                     'Gerbing 22.10.2015
    txtKommentar.height = Me.height - 1200                  'Gerbing 22.10.2015

    txtKommentar.Refresh                                    'Gerbing 16.09.2007
    On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    glngKommentarTop = Me.Top                               'Gerbing 30.05.2013
    glngKommentarLeft = Me.Left
    glngKommentarWidth = Me.width
    glngKommentarHeight = Me.height
    'Form1.KommentarFensterEinblenden = False                'Gerbing 19.07.2005                'Gerbing 28.10.2012
    On Error Resume Next
    'Me.Hide
    Unload KommentarForm                                    'Gerbing 23.10.2012
End Sub

