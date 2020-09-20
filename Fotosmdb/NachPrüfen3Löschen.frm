VERSION 5.00
Object = "{FCCB83BF-E483-4317-9FF2-A460758238B5}#1.2#0"; "CBLCtlsU.ocx"
Begin VB.Form NachPr�fen3L�schen 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Pr�fen3 - Die gefunden Dateien sind �berfl�ssig -> l�schen"
   ClientHeight    =   8028
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11928
   Icon            =   "NachPr�fen3L�schen.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8028
   ScaleWidth      =   11928
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton btnAlleMarkieren 
      Caption         =   "Alle markieren"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   7440
      Width           =   3492
   End
   Begin VB.CommandButton btnAbbrechen 
      Caption         =   "&Abbrechen"
      Height          =   375
      Left            =   8280
      TabIndex        =   1
      Top             =   7440
      Width           =   3492
   End
   Begin VB.CommandButton btnL�schen 
      Caption         =   "markierte Dateien &l�schen"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Zum Markieren k�nnen Sie die Tasten Umsch und Strg zu Hilfe nehmen"
      Top             =   7440
      Width           =   3492
   End
   Begin CBLCtlsLibUCtl.ListBox lstZus�tzlicheDateien 
      Height          =   7212
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   11652
      _cx             =   20553
      _cy             =   12721
      AllowDragDrop   =   0   'False
      AllowItemSelection=   -1  'True
      AlwaysShowVerticalScrollBar=   -1  'True
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   0
      ColumnWidth     =   -1
      DisabledEvents  =   1048811
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
      ScrollableWidth =   500
      Sorted          =   0   'False
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      ToolTips        =   0
      UseSystemFont   =   0   'False
      VirtualMode     =   0   'False
   End
End
Attribute VB_Name = "NachPr�fen3L�schen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim NL As String
    Dim Msg As String
    Public KollZus�tzlicheDateien As New Collection

Private Sub btnAbbrechen_Click()
        Do Until NachPr�fen3L�schen.KollZus�tzlicheDateien.Count = 0                    'Gerbing 03.11.2013
        NachPr�fen3L�schen.KollZus�tzlicheDateien.Remove 1
    Loop
    Do Until NachPr�fen3Aufnehmen.KollZus�tzlicheDateien.Count = 0                  'Gerbing 03.11.2013
        NachPr�fen3Aufnehmen.KollZus�tzlicheDateien.Remove 1
    Loop
    Me.Hide
End Sub

Private Sub btnAlleMarkieren_Click()
    Dim n As Long
    
    For n = 0 To lstZus�tzlicheDateien.ListItems.Count - 1
        lstZus�tzlicheDateien.ListItems(n).Selected = True
    Next n
End Sub

Private Sub btnL�schen_Click()
    Dim i As Long
    Dim n As Long
    Dim strTemp As String
    Dim MyAppPath As String
    Dim rc As Boolean
    Dim antwort As Long
    
    If lstZus�tzlicheDateien.ListItems.Count <> 0 Then                                          'Gerbing 26.10.2013
        Msg = LoadResString(Sprache + 2466)
        'msg = Wollen Sie wirklich die markierten Dateien l�schen (in den Papierkorb)?
        antwort = MsgBox(Msg, vbYesNo)
        If antwort = vbNo Then
            Exit Sub
        End If
    End If

    If gblnSQLServerVersion = True Then                     'Gerbing 05.09.2013
        MyAppPath = PublicLocationFotos
    Else
        MyAppPath = AppPath
    End If
    i = 0
    Screen.MousePointer = vbHourglass
    lstZus�tzlicheDateien.Visible = False                   'Gerbing 26.01.2006
    Do While lstZus�tzlicheDateien.ListItems.Count <> 0
        If lstZus�tzlicheDateien.ListItems(i).Selected Then
            strTemp = Replace(lstZus�tzlicheDateien.ListItems(i), "+:\", MyAppPath & "\")       'Gerbing 11.04.2005
            'Kill strTemp                                                                       'Gerbing 11.04.2005
            rc = file_delete(strTemp, True, True) '2.parameter True l�scht in den Papierkorb    'Gerbing 05.09.2013 26.10.2013
            lstZus�tzlicheDateien.ListItems.Remove i
            i = 0
        Else
            If i < lstZus�tzlicheDateien.ListItems.Count - 1 Then
                i = i + 1
            Else
                Exit Do
            End If
        End If
    Loop
    lstZus�tzlicheDateien.Visible = True                'Gerbing 26.01.2006
    Screen.MousePointer = vbDefault
    Form1.FehlerGefunden = False
    Form1.txtFehlerU.Text = ""
End Sub

Private Sub Form_Load()
    Call AnpassenNutzerWunsch(Me)                                       'Gerbing 11.03.2017
    Me.Caption = LoadResString(1343 + Sprache)  'Pr�fen3 - Die gefunden Dateien sind �berfl�ssig -> l�schen
    btnL�schen.Caption = LoadResString(1344 + Sprache)      'markierte Dateien &l�schen
    btnAbbrechen.Caption = LoadResString(3013 + Sprache)            '&Abbrechen
    btnL�schen.ToolTipText = LoadResString(1432 + Sprache)           'Zum Markieren k�nnen Sie die Tasten Umsch und Strg zu Hilfe nehmen
    btnAlleMarkieren.Caption = LoadResString(1518 + Sprache) 'Alle mar&kieren
    
    'lstZus�tzlicheDateien.MultiSelect = 2 muss in der Entwicklungsumgebung eingestellt werden
    If lstZus�tzlicheDateien.ListItems.Count <> 0 Then
        'lstZus�tzlicheDateien.ListIndex = 0
    End If
    NL = vbNewLine
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lstZus�tzlicheDateien.width = NachPr�fen3L�schen.width - 400
    lstZus�tzlicheDateien.height = NachPr�fen3L�schen.height - 1240
    btnL�schen.top = Me.height - 975
    btnAlleMarkieren.top = Me.height - 975
    btnAbbrechen.top = Me.height - 975
    On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    Me.Hide
End Sub
