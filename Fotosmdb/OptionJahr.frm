VERSION 5.00
Begin VB.Form frmOptionJahr 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Einstellung f�r das Feld 'Jahr'"
   ClientHeight    =   4212
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   10908
   Icon            =   "OptionJahr.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4212
   ScaleWidth      =   10908
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   3492
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10692
      Begin VB.TextBox txtDummy 
         Height          =   372
         Left            =   2760
         TabIndex        =   6
         Text            =   "9999"
         Top             =   2880
         Width           =   732
      End
      Begin VB.OptionButton optExif 
         Caption         =   "Die Jahreszahl soll aus dem gew�hlten EXIF/IPTC-Feld importiert werden"
         Height          =   372
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   10332
      End
      Begin VB.OptionButton optManuell 
         Caption         =   $"OptionJahr.frx":038A
         Height          =   732
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   10332
      End
      Begin VB.OptionButton optExtrahieren 
         Caption         =   "Das Jahr wird aus dem Dateinamen extrahiert"
         Height          =   372
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   10332
      End
      Begin VB.Label Label1 
         Caption         =   "Immer wenn der Computer keine Jahreszahl findet, wird eine Dummy-Jahreszahl benutzt, standardm��ig die Zahl 9999"
         Height          =   612
         Left            =   240
         TabIndex        =   7
         Top             =   2160
         Width           =   10332
      End
      Begin VB.Label lblDummy 
         Caption         =   "Dummy-Jahreszahl:"
         Height          =   372
         Left            =   240
         TabIndex        =   5
         Top             =   2880
         Width           =   2412
      End
   End
End
Attribute VB_Name = "frmOptionJahr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnOK_Click()
    If txtDummy.Text = "" Then
        'MsgBox "Sie m�ssen eine g�ltige Dummy-Jahreszahl angeben"
        MsgBox LoadResString(1493 + Sprache)
        Exit Sub
    End If
    If Not IsNumeric(txtDummy.Text) Then
        'MsgBox "Sie m�ssen eine g�ltige Dummy-Jahreszahl angeben"
        MsgBox LoadResString(1493 + Sprache)
        txtDummy.Text = ""
        Exit Sub
    End If
    NeueDatens�tzeGenerieren.DummyJahr = txtDummy
    Unload Me
End Sub

Private Sub Form_Load()
    Call AnpassenNutzerWunsch(Me)                                       'Gerbing 11.03.2017
    Me.Caption = LoadResString(1484 + Sprache)              'Einstellung f�r das Feld 'Jahr'
    optExtrahieren.Caption = LoadResString(1353 + Sprache)  'Das Jahr wird aus dem Dateinamen extrahiert
    optManuell.Caption = LoadResString(1485 + Sprache)   'Wenn das Jahr nicht aus dem Dateinamen extrahiert werden kann, will ich vor dem Computer sitzen bleiben und die richtige Jahreszahl auf Aufforderung manuell vergeben
    optExif.Caption = LoadResString(1486 + Sprache) 'Die Jahreszahl soll aus dem gew�hlten EXIF/IPTC-Feld importiert werden
    lblDummy.Caption = LoadResString(1488 + Sprache)    'Dummy-Jahreszahl:
    Label1.Caption = LoadResString(1489 + Sprache)   'Immer wenn der Computer keine Jahreszahl findet, wird eine Dummy-Jahreszahl benutzt, standardm��ig das aktuelle Jahr
    btnOK.Caption = LoadResString(3001 + Sprache)   '&OK
    
    NeueDatens�tzeGenerieren.blnOptGew�hlt = False
    If NeueDatens�tzeGenerieren.chkUnbeaufsichtigt.Value = 1 Then
        optManuell.Enabled = False
    End If
    If NeueDatens�tzeGenerieren.chkExif.Value = 0 Then
        optExif.Enabled = False
    End If
    
    If NeueDatens�tzeGenerieren.optExtrahieren = True Then
        optExtrahieren = True
    End If
    If NeueDatens�tzeGenerieren.optManuell = True Then
        optManuell = True
    End If
    If NeueDatens�tzeGenerieren.optExif = True Then
        optExif = True
    End If
'    If NeueDatens�tzeGenerieren.DummyJahr <> "" Then
'        txtDummy.Text = NeueDatens�tzeGenerieren.DummyJahr
'    End If
    txtDummy.Text = Year(Now)                                                       'Gerbing 22.12.2019
End Sub

Private Sub Form_Unload(Cancel As Integer)
    NeueDatens�tzeGenerieren.DummyJahr = txtDummy
End Sub

Private Sub optExif_Click()
'    If NeueDatens�tzeGenerieren.cmbJahrEx.Text = "" Then
'        'MsgBox "Sie m�ssen angegeben aus welchem EXIF-Feld das Jahr importiert werden soll und vergessen Sie nicht anschlie�end nochmals auf 'Jahr...' zu klicken"
'        MsgBox LoadResString(1496 + Sprache)
'        Unload Me
'        NeueDatens�tzeGenerieren.cmbJahrEx.SetFocus
'        Exit Sub
'    End If
    NeueDatens�tzeGenerieren.blnOptGew�hlt = True
    NeueDatens�tzeGenerieren.optExif = True
End Sub

Private Sub optExtrahieren_Click()
    Dim blnJahrGefunden As Boolean
    Dim start As Long
    Dim PosBackSlash As Long
    Dim PosNextBackSlash As Long
    Dim Teilstring As String
    
    NeueDatens�tzeGenerieren.blnOptGew�hlt = True
    '4-stellige Jahreszahl finden im ersten Satz des Drag&Drop Containers (List1)
    start = 1
    Do
        NeueDatens�tzeGenerieren.txtArbeitsfortschritt.Text = NeueDatens�tzeGenerieren.List1.ListItems(0).Text
        PosBackSlash = InStr(start, NeueDatens�tzeGenerieren.List1.ListItems(0).Text, "\")
        If PosBackSlash = 0 Then
            blnJahrGefunden = False
            Exit Do
        End If
        PosNextBackSlash = InStr(PosBackSlash + 1, NeueDatens�tzeGenerieren.List1.ListItems(0).Text, "\")
        If PosNextBackSlash = 0 Then
            blnJahrGefunden = False
            Exit Do
        Else
            'pr�fe den Teilstring ob es eine 4-stellige Jahreszahl ist
            Teilstring = Mid(NeueDatens�tzeGenerieren.List1.ListItems(0).Text, PosBackSlash + 1, PosNextBackSlash - PosBackSlash - 1)
            If Not IsNumeric(Teilstring) Or Len(Teilstring) <> 4 Then
                'es ist keine 4-stellige Jahreszahl
            Else
                'es ist eine 4-stellige Jahreszahl
                blnJahrGefunden = True
                Exit Do
            End If
        End If
        start = PosBackSlash + 1
    Loop
    'Wenn im ersten Satz des Drag&Drop Containers (List1) kein Jahr gefunden wird, muss eine MsgBox kommen
    'und es wird optDummy mit Jahreszahl 9999 eingeschaltet
    If blnJahrGefunden = False Then
        
        DoEvents
        'MsgBox "Der erste Dateiname im Drag&Drop-Container enth�lt keine Jahreszahl. Wenn Sie trotzdem diese Option benutzen, tr�gt das Programm eine 'Dummy-Jahreszahl' ein"
        MsgBox LoadResString(1495 + Sprache)
    End If
    NeueDatens�tzeGenerieren.optExtrahieren = True
End Sub

Private Sub optManuell_Click()
    NeueDatens�tzeGenerieren.blnOptGew�hlt = True
    NeueDatens�tzeGenerieren.optManuell = True
End Sub
