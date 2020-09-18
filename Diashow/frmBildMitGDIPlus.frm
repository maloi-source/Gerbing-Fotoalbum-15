VERSION 5.00
Begin VB.Form frmBildMitGDIPlus 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   4272
   ClientLeft      =   60
   ClientTop       =   24
   ClientWidth     =   10656
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   356
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   888
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   7  'Diagonalkreuz
      Height          =   852
      Left            =   120
      Top             =   120
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Menu mnuDatei 
      Caption         =   "&Datei..."
      Begin VB.Menu mnuDateiKopieren 
         Caption         =   "Datei kopieren..."
      End
      Begin VB.Menu mnuZeigeGeoPosition 
         Caption         =   "Zeige Geo-Position"
      End
      Begin VB.Menu mnuÖffneVerknüpfteAnwendung 
         Caption         =   "Öffne die mit der aktuellen Datei verknüpfte Anwendung"
      End
      Begin VB.Menu mnuÖffneDruckProgramm 
         Caption         =   "Öffne das Druckprogramm für die aktuelle Datei..."
      End
      Begin VB.Menu mnuEmailSenden 
         Caption         =   "Email mit Anhang senden..."
      End
      Begin VB.Menu mnuÖffneExploreFenster 
         Caption         =   "Öffne ein Explorer-Fenster, wo die aktuelle Datei markiert ist"
      End
      Begin VB.Menu mnuOptionen 
         Caption         =   "Optionen..."
      End
      Begin VB.Menu mnuVerschiebenPapierkorb 
         Caption         =   "Verschiebe alle mit Häkchen markierten Dateien in den Papierkorb"
      End
      Begin VB.Menu mnuKopiereHäkchen 
         Caption         =   "Kopiere alle mit Häkchen markierten Dateien"
      End
      Begin VB.Menu mnuÄndernDateinamen 
         Caption         =   "Ändern alle mit Häkchen markierten Dateinamen"
      End
   End
   Begin VB.Menu mnuHilfe 
      Caption         =   "&Hilfe"
   End
   Begin VB.Menu mnuBeenden 
      Caption         =   "&Benden"
   End
   Begin VB.Menu mnuVersion 
      Caption         =   "&Version"
   End
End
Attribute VB_Name = "frmBildMitGDIPlus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Private StartX As Long
    Private StartY As Long
    Private EndX As Long
    Private EndY As Long
    Private RX As Long
    Dim sngWidth As Single
    Dim sngHeight As Single
    Dim lngPointer As Long
    Dim x As Long
    Dim y As Long
    
    Private Enum RotateFlipType                                                                         'Gerbing 10.05.2019
        RotateNoneFlipNone = 0
        Rotate90FlipNone = 1
        Rotate180FlipNone = 2
        Rotate270FlipNone = 3
        RotateNoneFlipX = 4
        Rotate90FlipX = 5
        Rotate180FlipX = 6
        Rotate270FlipX = 7
        RotateNoneFlipY = Rotate180FlipX
        Rotate90FlipY = Rotate270FlipX
        Rotate180FlipY = RotateNoneFlipX
        Rotate270FlipY = Rotate90FlipX
        RotateNoneFlipXY = Rotate180FlipNone
        Rotate90FlipXY = Rotate270FlipNone
        Rotate180FlipXY = RotateNoneFlipNone
        Rotate270FlipXY = Rotate90FlipNone
    End Enum
    Private Declare Function GdipImageRotateFlip Lib "gdiplus" _
                            (ByVal image As Long, ByVal rfType As RotateFlipType) As Status             'Gerbing 10.05.2019


Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim ShiftDown As Boolean
    Dim AltDown As Boolean
    Dim CtrlDown As Boolean

    If KeyCode = 0 Then Exit Sub                                        'Gerbing 25.08.2013
    
    'Das Entladen der Form frmBildMitGDIPlus mit Unload Me und anschließende Neuladen aus DiashowForm heraus ist nötig,
    'weil ich keinen anderen Weg gefunden habe die Überreste eines gezeichneten Bildes zu löschen bevor ein neues Bild
    'gezeichnet wird. Wenn ich das nicht mache, übermalen neue Bilder schon gezeichnete Bilder.
    
    ListBoxForm.chkExifAnzeigen.Value = 0
    ListBoxForm.chkExifAnzeigen.Enabled = True
    ListBoxForm.chkIptcAnzeigen.Value = 0
    ListBoxForm.chkIptcAnzeigen.Enabled = True
    ListBoxForm.txtIPTCInfo.Visible = False
    ListBoxForm.txtEXIFInfo.Visible = False
    'Unload HilfeBoxForm                                                'sonst flackert es beim Zoomen
    HilfeBoxForm.Hide
    
    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
    
    If Shift = vbAltMask And KeyCode = 18 Then  '18 = Menu key          'Gerbing 29.01.2018
        KeyCode = 0
        mnuDatei.Visible = Not mnuDatei.Visible
        mnuBeenden.Visible = Not mnuBeenden.Visible
        mnuVersion.Visible = Not mnuVersion.Visible
        mnuHilfe.Visible = Not mnuHilfe.Visible
        Sleep 100  'Sonst flackert es bei Festhalten der Taste Alt
        DoEvents
        Exit Sub
    End If
    
    DiashowForm.F8Gedrückt = False
    If KeyCode = vbKeyF4 Then
        If Shift = vbAltMask Then                                       'Alt+F4 gleichzeitig
            Set DiashowForm.EXF = Nothing
            Call DiashowForm.MDIFormMitTitle
            Set fso = Nothing
            If DiashowForm.gblnWithTitle = False Then
                Call DiashowForm.MDIFormMitTitle
            End If
            Unload frmDurchsuchenUnterOrdner
            Unload HilfeBoxForm
            Unload frmBildMitGDIPlus
            Unload WertxForm
            Unload XYPos
            Unload ListBoxForm
            End
        End If
    End If

    If ListBoxForm.ExLVwU.ListItems.Count <> 0 Then
        Select Case KeyCode
            Case 40                                                     'Pfeil nach unten
                If Shift = vbShiftMask Then                             'Pfeil nach unten und gleichzeitig Shift-Taste
                    Unload Me
                    Call DiashowForm.Form_KeyDown(KeyCode, Shift)
                End If
            Case 38                                                     'Pfeil nach oben
                If Shift = vbShiftMask Then                             'Pfeil nach oben und gleichzeitig Shift-Taste
                    Unload Me
                    Call DiashowForm.Form_KeyDown(KeyCode, Shift)
                End If
            Case vbKeyLeft
                If Shift = vbShiftMask Then                             'Pfeil nach links und gleichzeitig Shift-Taste
                    Unload Me
                    Call DiashowForm.Form_KeyDown(KeyCode, Shift)
                End If
                If Shift = vbAltMask Then                               'Alt und Pfeil nach links wirkt wie F2
                    Unload Me
                    Call DiashowForm.Form_KeyDown(KeyCode, Shift)
                End If
            Case vbKeyRight
                If Shift = vbShiftMask Then                             'Pfeil nach rechts und gleichzeitig Shift-Taste
                    Unload Me
                    Call DiashowForm.Form_KeyDown(KeyCode, Shift)
                End If
                If Shift = vbAltMask Then                               'Alt und Pfeil nach rechts wirkt wie F3
                    Unload Me
                    Call DiashowForm.Form_KeyDown(KeyCode, Shift)
                End If
            Case 33                                                     'Bild nach oben geht zum Anfang
                Shift = 0
                Unload Me
                Call DiashowForm.Form_KeyDown(KeyCode, Shift)
            Case vbKeyHome                                              'Pos1-Taste geht zum Anfang
                Shift = 0
                Unload Me
                Call DiashowForm.Form_KeyDown(KeyCode, Shift)
            Case 34                                                     'Bild nach unten geht zum Ende
                Shift = 0
                Unload Me
                Call DiashowForm.Form_KeyDown(KeyCode, Shift)
            Case vbKeyEnd                                               'Ende-Taste geht zum Ende
                Shift = 0
                Unload Me
                Call DiashowForm.Form_KeyDown(KeyCode, Shift)
            Case vbKeyF1                                                'F1
                Shift = 0
                Unload Me
                'On Error Resume Next                                                        'Gerbing 30.05.2013
                Call DiashowForm.Form_KeyDown(KeyCode, Shift)
                'frmBildMitGDIPlus.Show
            Case vbKeyF2                                                'F2
                Shift = 0
                Unload Me
                Call DiashowForm.Form_KeyDown(KeyCode, Shift)
            Case vbKeyF3                                                'F3
                Shift = 0
                Unload Me
                Call DiashowForm.Form_KeyDown(KeyCode, Shift)
            Case vbKeyF4                                                'F4
                Shift = 0
                Unload Me
                'On Error Resume Next                                                        'Gerbing 30.05.2013
                Call DiashowForm.Form_KeyDown(KeyCode, Shift)
                'frmBildMitGDIPlus.Show
            Case vbKeyF5                                                'F5
                ListBoxForm.Caption = LoadResString(1238 + Sprache)                          'Gerbing 10.06.2013 "Dateinamen"
                ListBoxForm.Show
                ListBoxForm.ZOrder 0
                'ListBoxForm.chkIptcAnzeigen.Value = 1                                      'Gerbing 09.06.2014
            Case vbKeyF6                                                'F6
                DiashowForm.F6TimerGestartet = True
                DiashowForm.Timer1.Enabled = True
            Case vbKeyF7                                                'F7
                DiashowForm.F6TimerGestartet = False
                DiashowForm.Timer1.Enabled = False
            Case vbKeyF8                                                'F8
                Shift = 0
                Unload Me
                Call DiashowForm.Form_KeyDown(KeyCode, Shift)
            Case vbKeyF9                                                'F9
                If DiashowForm.gblnMouseSichtbar = True Then
                    DiashowForm.gblnMouseSichtbar = False
                    Me.MousePointer = vbCustom      '99
                'Me.MouseIcon = LoadPicture(AppPath & "\MOUSE01.ICO")                       'Gerbing 29.07.2007
                Me.MouseIcon = LoadResPicture(104, 1)                                       'Gerbing 04.03.2013
                Else
                    DiashowForm.gblnMouseSichtbar = True
                    If DiashowForm.RechteckLupeScharf = True Then
                        Me.MousePointer = vbCustom
                    'Me.MouseIcon = LoadPicture(AppPath & "\SquareZoom.ico")                'Gerbing 29.07.2007
                    Me.MouseIcon = LoadResPicture(105, 1)                           'Gerbing 04.03.2013
                    Else
                        Me.MousePointer = vbDefault
                    End If
                End If
            Case vbKeyF10                                               'F10
                Me.WindowState = 2  '2=maximized
                WertxForm.Show
                WertxForm.ZOrder 0
            Case vbKeyF11                                               'F11
                XYPos.Show
            Case vbKeyZ
                If Shift = vbCtrlMask Then                              'Strg+Z gleichzeitig
                    If DiashowForm.ZoomFaktor <> 0 Then
                        KeyCode = vbKeySeparator                        'vbKeySeparator als Kennzeichnung für Neuzeichnen anstelle Zoom
                        Shift = 0
                        Unload Me
                        Call DiashowForm.Form_KeyDown(KeyCode, Shift)
                    End If
                    Me.MousePointer = vbCustom
                    'Me.MouseIcon = LoadPicture(AppPath & "\SquareZoom.ico")                'Gerbing 29.07.2007
                    Me.MouseIcon = LoadResPicture(105, 1)                           'Gerbing 04.03.2013
                    DiashowForm.RechteckLupeScharf = True
                End If
            Case vbKeyO
                If Shift = vbCtrlMask Then                              'Strg+O gleichzeitig
                    Me.MousePointer = vbDefault
                    DiashowForm.RechteckLupeScharf = False
                End If
            Case vbKeyC                                                 'Strg+C gleichzeitig Gerbing 12.08.2017
                If Shift = vbCtrlMask Then
                    Me.MousePointer = vbDefault
                    frmZwischenablageOderOrdner.Show 1
                End If
            Case vbKeyG                                                 'Strg+G gleichzeitig Gerbing 29.01.2018
                If Shift = vbCtrlMask Then
                    Call ListBoxForm.ZeigeGeoPosition
                End If
            Case vbKeyE                                                 'Strg+E gleichzeitig Gerbing 22.07.2020
                If Shift = vbCtrlMask Then
                    ListBoxForm.ExLVwU.ListItems(ListBoxForm.ExLVwUIndex).StateImageIndex = 2   'einschalten
                End If
            Case vbKeyA                                                 'Strg+A gleichzeitig Gerbing 22.07.2020
                If Shift = vbCtrlMask Then
                    ListBoxForm.ExLVwU.ListItems(ListBoxForm.ExLVwUIndex).StateImageIndex = 1   'ausschalten
                End If
        End Select
    Else                                                                'Gerbing 12.03.2008
        Shift = 0
        Unload Me
        Call DiashowForm.Form_KeyDown(KeyCode, Shift)
        Call DiashowForm.EsIstF8
    End If
End Sub

Private Sub Form_Load()
    Dim udtData As GDIPlusStartupInput
    
    Me.WindowState = 2                          '2=maximiert
    'Me.ControlBox = False muss in der Entwicklungsumgebung gesetzt werden
    'Me.Caption = ListBoxForm.ExLVwU.ListItems(ListBoxForm.ExLVwUIndex)
    'Achtung in der IDE wird unicode in Form.Caption nicht angezeigt

    udtData.GdiplusVersion = 1
    If GdiplusStartup(m_lngInstance, udtData, 0) Then
        MsgBox "GDI+ could not be initialized", vbCritical
        Exit Sub
    End If
    If GdipCreateFromHDC(Me.hdc, m_lngGraphics) Then
        MsgBox "Graphics object could not be created", vbCritical
        Exit Sub
    End If
    If gblnListBoxNeuDblClick = True Then                                                               'Gerbing 25.08.2013
        Call MyDrawImage(gblStrAktuellGezeigtesBild, DiashowForm.ZoomProzent)
        formCaption frmBildMitGDIPlus.hwnd, gblStrAktuellGezeigtesBild                                  'Gerbing 25.06.2013
    Else
        Call MyDrawImage(ListBoxForm.ExLVwU.ListItems(ListBoxForm.ExLVwUIndex), DiashowForm.ZoomProzent)
        formCaption frmBildMitGDIPlus.hwnd, ListBoxForm.ExLVwU.ListItems(ListBoxForm.ExLVwUIndex)       'Gerbing 25.06.2013
    End If
    Me.Show                                                                                             'Gerbing 25.06.2013
    gblnListBoxNeuDblClick = False                                                                      'Gerbing 25.08.2013
    mnuDatei.Visible = False                                                                            'Gerbing 29.01.2018
    mnuBeenden.Visible = False
    mnuVersion.Visible = False
    mnuHilfe.Visible = False
    mnuDatei.Caption = LoadResString(3167 + Sprache)        'Datei...
    mnuBeenden.Caption = LoadResString(1040 + Sprache)      'Beenden     'Gerbing 14.11.2007
    mnuHilfe.Caption = LoadResString(1039 + Sprache)        'Hilfe
    mnuDateiKopieren.Caption = LoadResString(3168 + Sprache) 'Datei kopieren...
    mnuEmailSenden.Caption = LoadResString(3169 + Sprache)  'Email mit Anhang senden...
    mnuÖffneExploreFenster.Caption = LoadResString(3170 + Sprache) 'Öffne ein Explorer-Fenster, wo die aktuelle Datei markiert ist
    mnuZeigeGeoPosition.Caption = LoadResString(3173 + Sprache) 'Zeige Geo-Position
    mnuVerschiebenPapierkorb.Caption = LoadResString(3189 + Sprache) 'Verschiebe alle mit Häkchen markierten Dateien in den Papierkorb
    mnuÖffneDruckProgramm.Caption = LoadResString(3180 + Sprache)   'Öffne das Druckprogramm für die aktuelle Datei...
    mnuÖffneVerknüpfteAnwendung.Caption = LoadResString(3184 + Sprache) 'Öffne die mit der aktuellen Datei verknüpfte Anwendung
    mnuOptionen.Caption = LoadResString(3190 + Sprache)     'Optionen
    mnuÄndernDateinamen.Caption = LoadResString(3191 + Sprache) 'Ändern alle mit Häkchen markierten Dateinamen
    mnuKopiereHäkchen.Caption = LoadResString(3192 + Sprache)   'Kopiere alle mit Häkchen markierten Dateinamen
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If DiashowForm.RechteckLupeScharf = False And DiashowForm.blnListBoxNeuDblClick = False Then    'Gerbing 05.05.2005
        If DiashowForm.RechteckLupeScharf = False Then
            'Wenn die Rechteck-Lupe nicht scharf ist und auch kein Doppelklick vorliegt, wird mit der linken Maustaste das Bild verschoben
            'und zwar exakt an die mit der Maus angesteuerte Stelle
            If Button = vbLeftButton Then
                StartX = x
                StartY = y
                Me.MousePointer = vbCustom
                If WertxForm.CheckSpeichernBildPosition.Value = 1 Then
                    'Me.MouseIcon = LoadPicture(AppPath & "\FourArrowsSave.ico")    'Gerbing 29.07.2007
                    Me.MouseIcon = LoadResPicture(102, 1)                           'Gerbing 04.03.2013
                Else
                    'Me.MouseIcon = LoadPicture(AppPath & "\FourArrows.ico")        'Gerbing 29.07.2007
                    Me.MouseIcon = LoadResPicture(101, 1)                           'Gerbing 04.03.2013
                End If
                Exit Sub
            End If
        End If
    End If
    If Button = vbLeftButton Then
        'Auf dem Klickpunkt beginnt das Rechteck, anfangs mit Width und Height = 0
        'Me.Cls
        Shape1.Width = 0
        Shape1.Height = 0
        StartX = x
        StartY = y
        Shape1.Left = x
        Shape1.Top = y
        EndX = x
        EndY = y
        Shape1.Visible = True
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, MyX As Single, MyY As Single)
    Dim MylblXPos As Double                                                 'Gerbing 17.03.2017
    Dim MylblYPos As Double                                                 'Gerbing 17.03.2017
    
    If Button = vbLeftButton Then
        'es wird ein Rechteck gezeichnet, das stets das Breiten/Höhen-Verhältnis des Bildschirms SGVH einhält 'Gerbing 03.09.2010
        'es wird nur gezeichnet, wenn der Nutzer von links nach rechts und von oben nach unten zieht
        RX = MyX
        If MyX > EndX And MyY > EndY Then
            If ((MyX - StartX) * 4) <> ((MyY - StartY) * 3) Then
                RX = MyX + ((MyY - StartY) * DiashowForm.SGVH)
            End If
            If RX - StartX < 0 Then Exit Sub
            Shape1.Width = RX - StartX
            Shape1.Height = Shape1.Width / DiashowForm.SGVH
        End If
        EndX = MyX
        EndY = MyY
    End If
    If DiashowForm.ZoomProzent = 0 Then Exit Sub                            'Gerbing 24.09.2015
    MylblXPos = (MyX - x) / (DiashowForm.ZoomProzent / 100)                 'Gerbing 11.03.2017 bei / entsteht eine Zahl mit Komma
    XYPos.lblXPos = CLng(MylblXPos)                                         'Gerbing 17.03.2017 Konvertieren in Long
    If XYPos.lblXPos < 0 Or XYPos.lblXPos > gsngPicWidth Then
        XYPos.lblXPos = ""
    End If
    MylblYPos = (MyY - y) / (DiashowForm.ZoomProzent / 100)                 'Gerbing 11.03.2017 bei / entsteht eine Zahl mit Komma
    XYPos.lblYPos = CLng(MylblYPos)                                         'Gerbing 17.03.2017 Konvertieren in Long
    If XYPos.lblYPos < 0 Or XYPos.lblYPos > gsngPicHeight Then
        XYPos.lblYPos = ""
    End If
    XYPos.lblBildgröße = gsngPicWidth & " x " & gsngPicHeight
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim KeyCode As Integer
    Dim msg As String
    
    If DiashowForm.RechteckLupeScharf = False And DiashowForm.blnListBoxNeuDblClick = False Then    'Gerbing 05.05.2005
        'Wenn die Rechteck-Lupe nicht scharf ist und auch kein Doppelklick vorliegt, wird mit der linken Maustaste das Bild verschoben
        'und zwar exakt an die mit der Maus angesteuerte Stelle
        If Button = vbLeftButton And DiashowForm.RechteckLupeScharf = False Then
            'aber nur wenn überhaupt eine Verschiebung stattgefunden hat
            If StartX <> x And StartY <> y Then
                DiashowForm.glngDiffX = DiashowForm.glngDiffX + x - StartX
                DiashowForm.glngDiffY = DiashowForm.glngDiffY + y - StartY
                Call DiashowForm.SpeichernInBildPosList
                Me.MousePointer = vbDefault
                KeyCode = vbKeySeparator                                    'vbKeySeparator als Kennzeichnung für Verschiebung
                Shift = 0
                Unload Me
                Call DiashowForm.Form_KeyDown(KeyCode, Shift)
            End If
        End If
    End If
    DiashowForm.blnListBoxNeuDblClick = False
    
    If Button = vbRightButton And Shape1.Visible = False Then
        'Me.WindowState = 2  '2=maximized
        HilfeBoxForm.Show
        Exit Sub
    End If

    If Shape1.Visible = False Then Exit Sub
    
    If DiashowForm.RechteckLupeScharf = True Then
        If EndX = StartX Or EndY = StartY Then
            Shape1.Visible = False
            'Wenn der Nutzer nur ins Bild klickt, bekommt er eine MsgBox
            'Wenn der Nutzer nur ins Bild klickt, bekommt er eine MsgBox
            msg = "Wenn Sie die Rechteck-Lupe benutzen wollen, müssen Sie mit der Maus ein Rechteck zeichnen." & vbNewLine
            'msg = LoadResString(2049 + Sprache) & vbNewLine
            msg = msg & "Sie haben nur in das Bild geklickt." & vbNewLine & vbNewLine
            'msg = msg & LoadResString(2050 + Sprache) & vbNewLine & vbNewLine
    
            msg = msg & "Mit Rechtsklick auf ein schon gezeichnetes Rechteck, können Sie die Zoom-Ansicht wiederholen"
            'msg = msg & LoadResString(2051 + Sprache)
            MsgBox msg
            'ImageForm.Image1.Refresh                          'Gerbing 06.06.2003
            Exit Sub
        End If
        'wenn nicht ein Rechteck mit einer mindestgröße vorliegt, wird FormZoom nicht geöffnet
        If (EndX - StartX) < 40 Or (EndY - StartY) < 30 Then
            msg = "Die Rechteck-Lupe ist scharf." & vbNewLine
            'msg = LoadResString(2052 + Sprache) & vbNewLine
            msg = msg & "Zu kleine Bildausschnitte werden nicht gezoomt." & vbNewLine
            'msg = msg & LoadResString(2053 + Sprache) & vbNewLine
            msg = msg & "Möglicherweise haben sie nur ins Bild geklickt" & vbNewLine
            'msg = msg & LoadResString(2054 + Sprache) & vbNewLine
            msg = msg & "oder nicht von links nach rechts und von oben nach unten gezogen."
            'msg = msg & LoadResString(2055 + Sprache)
            MsgBox msg
            Exit Sub
        End If
        DiashowForm.DifferenzX = DiashowForm.glngSaveX - StartX
        DiashowForm.DifferenzY = DiashowForm.glngSaveY - StartY
        DiashowForm.ZoomFaktor = DiashowForm.screenWidth / Shape1.Width
        KeyCode = vbKeySeparator                                    'vbKeySeparator als Kennzeichnung für Rechtecklupe
        Shift = 0
        Unload Me
        Call DiashowForm.Form_KeyDown(KeyCode, Shift)
    End If
End Sub

Private Sub Form_Resize()
    'MsgBox "frmBildMitGDIPlus_Form_Resize"
    Me.WindowState = 2          '2=maximized
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If m_lngGraphics Then
        If GdipDeleteGraphics(m_lngGraphics) Then _
            MsgBox "Graphics object could not be deleted", vbCritical
    End If
    GdiplusShutdown m_lngInstance
    XYPos.Hide
    WertxForm.Hide
    ListBoxForm.Hide
    'HilfeBoxForm.Hide  'HilfeBoxForm auskommentiert sonst flackert das Bild weil HilfeBoxForm kurz sichtbar wird
    DiashowForm.Hide
End Sub

Public Sub MyDrawImage(FileName As String, ZoomPercent As Long)
    
    gstrEXIFOrientation = ""                                                                        'Gerbing 10.05.2019
    Call EXIFLesenOrientation                                                                       'Gerbing 10.05.2019
    GdipLoadImageFromFile StrPtr(FileName), lngPointer
    
    If gstrEXIFOrientation = "8" Then GdipImageRotateFlip lngPointer, Rotate270FlipNone             'Rotate270FlipNone = 3  Gerbing 30.09.2019
    If gstrEXIFOrientation = "7" Then GdipImageRotateFlip lngPointer, Rotate270FlipX                'Rotate270FlipX = 7     Gerbing 30.09.2019
    If gstrEXIFOrientation = "6" Then GdipImageRotateFlip lngPointer, Rotate90FlipNone              'Rotate90FlipNone = 1   Gerbing 30.09.2019
    If gstrEXIFOrientation = "5" Then GdipImageRotateFlip lngPointer, Rotate90FlipX                 'Rotate90FlipX = 5      Gerbing 30.09.2019
    If gstrEXIFOrientation = "4" Then GdipImageRotateFlip lngPointer, Rotate180FlipX                'Rotate180FlipX = 6     Gerbing 30.09.2019
    If gstrEXIFOrientation = "3" Then GdipImageRotateFlip lngPointer, Rotate180FlipNone             'Rotate180FlipNone = 2  Gerbing 30.09.2019
    If gstrEXIFOrientation = "2" Then GdipImageRotateFlip lngPointer, RotateNoneFlipX               'RotateNoneFlipX = 4    Gerbing 30.09.2019
    
    GdipGetImageDimension lngPointer, sngWidth, sngHeight
    'X und Y ausrechnen
    x = (DiashowForm.screenWidth \ 2) - ((ZoomPercent / 100) * (sngWidth \ 2))
    y = (DiashowForm.screenHeight \ 2) - ((ZoomPercent / 100) * (sngHeight \ 2))
    x = x + DiashowForm.glngDiffX
    y = y + DiashowForm.glngDiffY
    If DiashowForm.RechteckLupeScharf = True Then
        'Bei Rechteck-Zoom ist uninteressant was als X oder Y bisher errechnet wurde
        x = DiashowForm.DifferenzX * DiashowForm.ZoomFaktor
        y = DiashowForm.DifferenzY * DiashowForm.ZoomFaktor
        ZoomPercent = ZoomPercent * DiashowForm.ZoomFaktor
        Shape1.Visible = False
        DiashowForm.glngDiffX = 0
        DiashowForm.glngDiffY = 0
        DiashowForm.ZoomProzent = 100
        DiashowForm.RechteckLupeScharf = False      'Der Rechteck-Zoom funktioniert nur einmal
    End If
    DiashowForm.glngSaveX = x
    DiashowForm.glngSaveY = y
    ' Setzen der Optimierungsmodis
    GdipSetSmoothingMode m_lngGraphics, SmoothingModeNone
    GdipSetInterpolationMode m_lngGraphics, InterpolationModeHighQualityBicubic
    GdipSetPixelOffsetMode m_lngGraphics, PixelOffsetModeNone
    GdipSetCompositingQuality m_lngGraphics, CompositingQualityDefault
    GdipSetCompositingMode m_lngGraphics, CompositingModeSourceOver
    ' zoomen
    GdipDrawImageRect m_lngGraphics, lngPointer, x, y, sngWidth * ZoomPercent \ 100, sngHeight * ZoomPercent \ 100
    GdipDisposeImage lngPointer
    Me.Refresh
End Sub

Private Sub mnuÄndernDateinamen_Click()
    Call ListBoxForm.DateinamenÄndern                                                       'Gerbing 28.01.2018 29.01.2018
End Sub

Private Sub mnuBeenden_Click()
    End
End Sub

Private Sub mnuDateiKopieren_Click()
    frmZwischenablageOderOrdner.Show 1                                                      'Gerbing 29.01.2018
End Sub

Private Sub mnuEmailSenden_Click()
    Call ListBoxForm.EmailMitAnhangSenden                                                                        'Gerbing 13.08.2017 29.01.2018
End Sub

Private Sub mnuHilfe_Click()
    Dim retval As Long
    Dim DateinamenErweiterung As String
    Dim CHMFile As String
    Dim ErrorText As String
    Dim msg As String

    If Sprache = 0 Then
        'CommonDialog1.HelpFile = AppPath & "\help\deutsch\diashow.HLP"
        CHMFile = AppPath & "\Help\Deutsch\diashow.CHM"                         'Gerbing 14.03.2007
        retval = ShellExecute(Me.hwnd, "open", CHMFile, vbNull, vbNull, 1)
        If retval <= 32 Then
            DateinamenErweiterung = "CHM"
            ErrorText = GetShellError(retval)           'Gerbing 20.08.2008
            msg = "Errortext=" & ErrorText & vbNewLine
            msg = msg & "Errornr=" & retval & vbNewLine & vbNewLine
            
            msg = msg & CHMFile & vbNewLine
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
            msg = msg & LoadResString(2013 + Sprache)
            'MsgBox msg
            MessageBoxW 0, StrPtr(msg), StrPtr("GERBING Diashow"), vbInformation
        End If

    Else
        'CommonDialog1.HelpFile = AppPath & "\help\english\diashow.HLP"
        CHMFile = AppPath & "\Help\English\diashow.CHM"                           'Gerbing 14.03.2007
        retval = ShellExecute(Me.hwnd, "open", CHMFile, vbNull, vbNull, 1)
        If retval <= 32 Then
            DateinamenErweiterung = "CHM"
            ErrorText = GetShellError(retval)           'Gerbing 20.08.2008
            msg = "Errortext=" & ErrorText & vbNewLine
            msg = msg & "Errornr=" & retval & vbNewLine & vbNewLine
            
            msg = msg & CHMFile & vbNewLine
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
            msg = msg & LoadResString(2013 + Sprache)
            'MsgBox msg
            MessageBoxW 0, StrPtr(msg), StrPtr("GERBING Diashow"), vbInformation
        End If
    End If
End Sub

Private Sub mnuKopiereHäkchen_Click()
    Call ListBoxForm.KopierenMarkierte                                                      'Gerbing 29.01.2018
End Sub

Private Sub mnuÖffneDruckProgramm_Click()
    Shell ("rundll32.exe SHELL32,OpenAs_RunDLL " & ListBoxForm.ExLVwU.ListItems(ListBoxForm.ExLVwUIndex))                                'Gerbing 07.08.2013
End Sub

Private Sub mnuÖffneExploreFenster_Click()
    Dim retval As Long
    
    retval = RunShellExecute(Me.hwnd, "open", "explorer.exe", "/e,/select," & ListBoxForm.ExLVwU.ListItems(ListBoxForm.ExLVwUIndex), vbNull, 1) 'Gerbing 29.01.2018
End Sub

Private Sub mnuÖffneVerknüpfteAnwendung_Click()
    Call ListBoxForm.ÖffneVerknüpfteAnwendung                                   'Gerbing 29.01.2018
End Sub

Private Sub mnuOptionen_Click()                                                 'Gerbing 29.01.2018
    Me.WindowState = 2  '2=maximized
    WertxForm.Show
    WertxForm.ZOrder 0
End Sub

Private Sub mnuVerschiebenPapierkorb_Click()
    Call ListBoxForm.LöschenMarkierte                                                       'Gerbing 29.01.2018
End Sub

Private Sub mnuVersion_Click()
    Dim msg As String
    
    'Versions-Informationen ermitteln                                           'Gerbing 28.08.2013
    msg = "Version " & GetDiashowExeVersion                                     'Gerbing 28.08.2013
    MessageBoxW 0, StrPtr(msg), StrPtr("GERBING Diashow"), vbInformation
End Sub

Private Sub mnuZeigeGeoPosition_Click()
    Call ListBoxForm.ZeigeGeoPosition                                           'Gerbing 03.10.2016
End Sub
