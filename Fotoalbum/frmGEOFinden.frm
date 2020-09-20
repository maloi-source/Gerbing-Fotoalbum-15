VERSION 5.00
Begin VB.Form frmGEOFinden 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "GEO-Position finden"
   ClientHeight    =   7608
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   7944
   Icon            =   "frmGEOFinden.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   634
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   662
   StartUpPosition =   1  'Fenstermitte
   Begin VB.PictureBox Rechteck 
      BackColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   120
      ScaleHeight     =   444
      ScaleWidth      =   564
      TabIndex        =   5
      Top             =   5520
      Visible         =   0   'False
      Width           =   612
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
      TabIndex        =   3
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
      Left            =   120
      Style           =   2  'Dropdown-Liste
      TabIndex        =   2
      Top             =   120
      Width           =   2532
   End
   Begin VB.TextBox txtCenter 
      BackColor       =   &H8000000F&
      Height          =   288
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2532
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1500
      Left            =   960
      Max             =   20
      TabIndex        =   0
      Top             =   1800
      Value           =   16
      Visible         =   0   'False
      Width           =   252
   End
   Begin FotoAlbum.ucGMap ucGMap1 
      Height          =   7680
      Left            =   480
      TabIndex        =   4
      ToolTipText     =   "Ziehen Sie ein Rechteck"
      Top             =   0
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   13547
   End
   Begin VB.Image imgEarth 
      Height          =   252
      Left            =   120
      Picture         =   "frmGEOFinden.frx":038A
      Stretch         =   -1  'True
      ToolTipText     =   "Nach dem Klicken auf den Erdball können Sie ein Rechteck zeichnen"
      Top             =   1680
      Width           =   252
   End
   Begin VB.Image imgMinus 
      Height          =   432
      Left            =   0
      Picture         =   "frmGEOFinden.frx":07CC
      Stretch         =   -1  'True
      Top             =   840
      Width           =   432
   End
   Begin VB.Image imgPlus 
      Height          =   432
      Left            =   0
      Picture         =   "frmGEOFinden.frx":0B56
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   432
   End
End
Attribute VB_Name = "frmGEOFinden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'Ich will vor dem Rechteck zeichnen erstnoch den Ausschnitt verschieben oder zentrieren können, das regelt blnBeginneRechteck
    'Ich darf nur nur einmal ein Rechteck ziehen und muss dann die Form entladen,
    'sonst stimmt das Rechteck und die angegebenen Positionen nicht überein, mit anderen Worten das Rechteck liefert falsche Positionen
    
    Dim MarkerChar As String
    Dim Startpunkt As String
    Dim Endpunkt As String
    Private StartX As Long
    Private StartY As Long
    Private EndX As Long
    Private EndY As Long
    Private RX As Long
    Dim AbsstartX As Long
    Dim AbsstartY As Long
    Dim StartViertel As Long
    Dim EndViertel As Long
    Dim DebugStartX As Long
    Dim DebugStartY As Long
    Dim blnMouseDownOccured As Boolean
    Dim rstErsterStart As ADODB.Recordset                           'Gerbing 05.09.2016
    Dim SQL As String
    Public blnBeginneRechteck As Boolean

Private Sub Form_Load()
    Dim i As Long
    Dim GMouseCoordLatLng As String
    Dim GeoPosition As String

    Call AnpassenNutzerWunsch(Me)                                   'Gerbing 11.03.2017
    blnBeginneRechteck = False
    Me.ScaleMode = 3                                                '3=Pixel ist nötig damit das Rechteck korrekt gezeichnet wird
    'Me.Caption = "GEO-Position finden"                             'Gerbing 05.09.2016
    Me.Caption = LoadResString(3163 + Sprache)
    'ucGMap1.tooltipText = "Vergrößern Sie die Landkarte auf den gewünschten Ausschnitt, dann ziehen Sie mit der Maus ein Rechteck von links oben nach rechts unten"
    ucGMap1.tooltipText = LoadResString(2343 + Sprache)
    imgEarth.tooltipText = LoadResString(2347 + Sprache) 'Nach dem Klicken auf den Erdball können Sie ein Rechteck zeichnen
    For i = 0 To 20: cmbZoom.AddItem 2 ^ i: Next
    For i = 0 To 3: cmbMapType.AddItem ucGMap1.GetMapType(i): Next
    cmbMapType.ListIndex = 0
    'GeoPosition = "A: 50.83266,12.81863"  'Chemnitz Georg-Weerth-Straße
    'GeoPosition = "A: 40.44,-111.54"  'USA
    'GeoPosition = "50.83266,12.81863"  'Chemnitz Georg-Weerth-Straße
    'ich will den zuletzt benutzten Eckpunkt als Geoposition und als GMouseCoordLatLng anbieten, falls es einen gibt
    'Der steht in Tabelle ErsterStart Feld LetzterGEOPunkt                              'Gerbing 05.09.2016
    'Seit Version 14.2.2 gibt es in der Tabelle ErsterStart das Feld LetzterGEOPunkt und ZoomListIndex        'Gerbing 05.09.2016
    'Und in der Tabelle fotos die Felder GPSLatitude und GPSLongitude
    On Error Resume Next
    SQL = "select * From ErsterStart;"
    Set rstErsterStart = New ADODB.Recordset
    With rstErsterStart
        .Source = SQL
        .ActiveConnection = DBado                                                       'Gerbing 23.11.2017
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    If Err.Number <> 0 Then
        'hier existiert das Feld LetzterGEOPunkt nicht
        If gblnSchreibgeschützt = False Then
            If gblnSQLServerVersion = True Then
                'SQL Server
                DBado.Execute _
                    "ALTER TABLE ErsterStart ADD LetzterGEOPunkt VARCHAR(255)"          'es heißt ADD und nicht ADD COLUMN
                DBado.Execute _
                    "ALTER TABLE ErsterStart ADD ZoomListIndex VARCHAR(255)"
            Else
                'Access Version
                'also wird Feld LetzterGEOPunkt und ZoomListIndex erzeugt
                DBado.Execute _
                    "ALTER TABLE ErsterStart ADD COLUMN LetzterGEOPunkt TEXT"
                DBado.Execute _
                    "ALTER TABLE ErsterStart ADD COLUMN ZoomListIndex TEXT"
            End If
        End If
        'hier ist Spalte Geoposition nicht vorhanden ich benutze den Standort Rabenstein
        GeoPosition = "50.83266,12.81863"  'Chemnitz Oberfrohnaer Straße Burg Rabenstein
        GMouseCoordLatLng = "50.83266,12.81863"
    Else
        'hier existiert das Feld LetzterGEOPunkt und ZoomListIndex
        If rstErsterStart.Fields("LetzterGEOPunkt") = "" Or IsNull(rstErsterStart.Fields("LetzterGEOPunkt")) Then
            'hier ist LetzterGEOPunkt leer ich benutze den Standort Rabenstein
            GeoPosition = "50.83266,12.81863"  'Chemnitz Oberfrohnaer Straße Burg Rabenstein
            GMouseCoordLatLng = "50.83266,12.81863"
        Else
            GeoPosition = rstErsterStart.Fields("LetzterGEOPunkt")
            GMouseCoordLatLng = rstErsterStart.Fields("LetzterGEOPunkt")
        End If
        If rstErsterStart.Fields("ZoomListIndex") = "" Or IsNull(rstErsterStart.Fields("ZoomListIndex")) Then
            cmbZoom.ListIndex = 0                                                       'bei 0 da kommt die Erde in Gesamtansicht
            VScroll1.Value = 0                                                          'bei 0 da kommt die Erde in Gesamtansicht
        Else
            cmbZoom.ListIndex = (rstErsterStart.Fields("ZoomListIndex"))
            VScroll1.Value = (rstErsterStart.Fields("ZoomListIndex"))
        End If
    End If
    rstErsterStart.Close
    On Error GoTo 0
    'GMouseCoordLatLng = "50.83266,12.81863"
    MarkerChar = Chr$(65)
    ucGMap1.AddMarker GMouseCoordLatLng, vbGreen, MarkerChar
    
    ucGMap1.GPoint = Mid$(GeoPosition, 1)
    txtCenter.Text = "CenterPoint: " & ucGMap1.GPoint
    'txtMouseLatLng.Text = "MousePoint: " & ""
    ucGMap1.Refresh
    Me.Refresh
End Sub
 
Private Sub cmbZoom_Click()
    ucGMap1.GZoom = cmbZoom.ListIndex
End Sub

Private Sub cmbMapType_Click()
    ucGMap1.MapType = cmbMapType.ListIndex
End Sub

Private Sub imgEarth_Click()
    blnBeginneRechteck = True
    imgEarth.Visible = False
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

Private Sub ucGMap1_DblClick(ByVal GMouseCoordLatLng As String)
    ucGMap1.GPoint = Mid$(GMouseCoordLatLng, 1)                         'Gerbing 02.09.2016 auf Doppelklick wird zentriert
End Sub

Private Sub ucGMap1_MouseDown(ByVal GMouseCoordLatLng As String, Button As Integer, Shift As Integer, x As Single, y As Single)
    'X und Y werden als Abstände zum Karten-Mittelpunkt (AddMarker) angegeben
    'beide Werte sind negativ im Viertel1
    'beide Werte sind positiv im Viertel4
    
    If blnBeginneRechteck = False Then Exit Sub
    If Button = vbLeftButton Then
        DebugStartX = x
        DebugStartY = y
        StartViertel = 4
        AbsstartX = Abs(x)
        AbsstartY = Abs(y)
        If x < 0 And y < 0 Then
            StartViertel = 1
            StartX = (ucGMap1.width / 2) - Abs(x)
            StartY = (ucGMap1.height / 2) - Abs(y)
        End If
        If x >= 0 And y < 0 Then
            StartViertel = 2
            StartX = (ucGMap1.width / 2) + x
            StartY = (ucGMap1.height / 2) - Abs(y)
        End If
        If x < 0 And y >= 0 Then
            StartViertel = 3
            StartX = (ucGMap1.width / 2) - Abs(x)
            StartY = (ucGMap1.height / 2) + y
        End If
        If StartViertel = 4 Then
            StartX = (ucGMap1.width / 2) + x
            StartY = (ucGMap1.height / 2) + y
        End If
        Rechteck.Left = StartX + ucGMap1.Left
        Rechteck.Top = StartY
        Rechteck.width = 1
        Rechteck.height = 1
        Startpunkt = GMouseCoordLatLng
        blnMouseDownOccured = True
        Rechteck.Visible = False
    End If
End Sub

Private Sub ucGMap1_MouseMove(ByVal GMouseCoordLatLng As String, Button As Integer, Shift As Integer, x As Single, y As Single)
    'X und Y werden als Abstände zum Karten-Mittelpunkt (AddMarker) angegeben
    'beide Werte sind negativ im Viertel1
    'beide Werte sind positiv im Viertel4

    If blnBeginneRechteck = False Then Exit Sub
    'txtMouseLatLng.Text = "MousePoint: " & GMouseCoordLatLng
    If blnMouseDownOccured = False Then Exit Sub
    On Error GoTo Fehlermeldung
'    Debug.Print "MouseUpX=" & x
'    Debug.Print "MouseUpY=" & y
    Endpunkt = GMouseCoordLatLng
    EndViertel = 4
    If x < 0 And y < 0 Then EndViertel = 1
    If x >= 0 And y < 0 Then EndViertel = 2
    If x < 0 And y >= 0 Then EndViertel = 3

    If StartViertel = 1 And EndViertel = 1 Then
        Rechteck.width = AbsstartX - Abs(x)
        Rechteck.height = AbsstartY - Abs(y)
    End If
    If StartViertel = 1 And EndViertel = 2 Then
        Rechteck.width = AbsstartX + x
        Rechteck.height = AbsstartY - Abs(y)
    End If
    If StartViertel = 1 And EndViertel = 3 Then
        Rechteck.width = AbsstartX - Abs(x)
        Rechteck.height = AbsstartY + y
    End If
    If StartViertel = 1 And EndViertel = 4 Then
        Rechteck.width = AbsstartX + x
        Rechteck.height = AbsstartY + y
    End If
    If (StartViertel = 2 And EndViertel = 1) Or (StartViertel = 2 And EndViertel = 3) Or (StartViertel = 3 And EndViertel = 1) _
    Or (StartViertel = 3 And EndViertel = 2) Or (StartViertel = 4 And EndViertel = 1) Or (StartViertel = 4 And EndViertel = 2) _
    Or (StartViertel = 4 And EndViertel = 3) Then
        Rechteck.Visible = False
        MsgBox "Sie müssen ein Rechteck von oben links nach unten rechts ziehen"
        blnMouseDownOccured = False
        Exit Sub
    End If
    If StartViertel = 2 And EndViertel = 2 Then
        Rechteck.width = x - AbsstartX
        Rechteck.height = AbsstartY - Abs(y)
    End If
    If StartViertel = 2 And EndViertel = 4 Then
        Rechteck.width = x - AbsstartX
        Rechteck.height = AbsstartY + y
    End If
    If StartViertel = 3 And EndViertel = 3 Then
        Rechteck.width = AbsstartX - Abs(x)
        Rechteck.height = y - AbsstartY
    End If
        If StartViertel = 3 And EndViertel = 4 Then
        Rechteck.width = AbsstartX + x
        Rechteck.height = y - AbsstartY
    End If
    If StartViertel = 4 And EndViertel = 4 Then
        Rechteck.width = x - AbsstartX
        Rechteck.height = y - AbsstartY
    End If
    Rechteck.Visible = True
    Exit Sub
Fehlermeldung:
    On Error GoTo 0
    Rechteck.Visible = False
    MsgBox "Sie müssen ein Rechteck von oben links nach unten rechts ziehen"
    blnMouseDownOccured = False
End Sub

Private Sub ucGMap1_MouseUp(ByVal GMouseCoordLatLng As String, Button As Integer, Shift As Integer, x As Single, y As Single)
    'X und Y werden als Abstände zum Karten-Mittelpunkt (AddMarker) angegeben
    'beide Werte sind negativ im Viertel1
    'beide Werte sind positiv im Viertel4

    If blnBeginneRechteck = False Then Exit Sub
    blnMouseDownOccured = False
    If DebugStartX = x And DebugStartY = y Then
        Exit Sub
    End If
    Rechteck.Visible = True
    'MsgBox "Startpunkt=" & Startpunkt & vbNewLine & "Endpunkt=" & Endpunkt
    MsgBox LoadResString(2345 + Sprache) & Startpunkt & vbNewLine & LoadResString(2346 + Sprache) & Endpunkt
    gstrGEOStartPunkt = Startpunkt
    gstrGEOEndPunkt = Endpunkt
    
    'Jetzt werden LetzterGEOPunkt und ZoomListIndex in der Tabelle ErsterStart gespeichert
    On Error Resume Next
    SQL = "select * From ErsterStart;"
    Set rstErsterStart = New ADODB.Recordset
    With rstErsterStart
        .Source = SQL
        .ActiveConnection = DBado                                                   'Gerbing 23.11.2017
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    If Err.Number = 0 Then
        rstErsterStart.Fields("LetzterGEOPunkt") = Startpunkt
        rstErsterStart.Fields("ZoomListIndex") = cmbZoom.ListIndex
        rstErsterStart.Update
    End If
    rstErsterStart.Close
    Unload Me
End Sub

Private Sub VScroll1_Change()
    cmbZoom.ListIndex = VScroll1.Value
End Sub


