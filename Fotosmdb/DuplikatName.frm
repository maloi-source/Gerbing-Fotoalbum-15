VERSION 5.00
Begin VB.Form DuplikatName 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Fehler Doppelter Name"
   ClientHeight    =   3480
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8448
   ControlBox      =   0   'False
   Icon            =   "DuplikatName.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   8448
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton btnAbbrechen 
      Caption         =   "&Abbrechen"
      Height          =   495
      Left            =   5040
      TabIndex        =   6
      Top             =   2760
      Width           =   3372
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&Geänderten Name benutzen"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   3372
   End
   Begin VB.TextBox txtManuellerName 
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   8295
   End
   Begin VB.Label Label3 
      Caption         =   "Sie dürfen nur den Dateinamen-Anteil ändern, der nach dem Jahr steht"
      Height          =   372
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   8292
   End
   Begin VB.Label Label2 
      Caption         =   "Ändern Sie jetzt den Dateinamen, so daß kein Duplikat entsteht."
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   8172
   End
   Begin VB.Label lblDuplikatName 
      BorderStyle     =   1  'Fest Einfach
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   8292
   End
   Begin VB.Label Label1 
      Caption         =   "Im Zielordner gibt es bereits eine Datei mit einem gleichen Namen"
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8172
   End
End
Attribute VB_Name = "DuplikatName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim VordemJahr As String
    Dim JahrInFilename As String
    Dim NachDemJahr As String
    Dim Pos As Long

Private Sub btnAbbrechen_Click()
    gblnAbbrechen = True
    Unload Me
End Sub

Private Sub btnOK_Click()
    Dim temp As String
    
    'Prüfen, ob verbotene Zeichen vorkommen
    Pos = InStr(txtManuellerName, "\")
    If Pos <> 0 Then
        'MsgBox "Ein Dateiname darf keines der folgenden Zeichen enthalten: \/:*?""<>|"
        MsgBox LoadResString(1367 + Sprache)
        Exit Sub
    End If
    Pos = InStr(txtManuellerName, "/")
    If Pos <> 0 Then
        'MsgBox "Ein Dateiname darf keines der folgenden Zeichen enthalten: \/:*?""<>|"
        MsgBox LoadResString(1367 + Sprache)
        Exit Sub
    End If
    Pos = InStr(txtManuellerName, ":")
    If Pos <> 0 Then
        'MsgBox "Ein Dateiname darf keines der folgenden Zeichen enthalten: \/:*?""<>|"
        MsgBox LoadResString(1367 + Sprache)
        Exit Sub
    End If
    Pos = InStr(txtManuellerName, "*")
    If Pos <> 0 Then
        'MsgBox "Ein Dateiname darf keines der folgenden Zeichen enthalten: \/:*?""<>|"
        MsgBox LoadResString(1367 + Sprache)
        Exit Sub
    End If
    Pos = InStr(txtManuellerName, "?")
    If Pos <> 0 Then
        'MsgBox "Ein Dateiname darf keines der folgenden Zeichen enthalten: \/:*?""<>|"
        MsgBox LoadResString(1367 + Sprache)
        Exit Sub
    End If
    Pos = InStr(txtManuellerName, """")
    If Pos <> 0 Then
        'MsgBox "Ein Dateiname darf keines der folgenden Zeichen enthalten: \/:*?""<>|"
        MsgBox LoadResString(1367 + Sprache)
        Exit Sub
    End If
    Pos = InStr(txtManuellerName, "<")
    If Pos <> 0 Then
        'MsgBox "Ein Dateiname darf keines der folgenden Zeichen enthalten: \/:*?""<>|"
        MsgBox LoadResString(1367 + Sprache)
        Exit Sub
    End If
    Pos = InStr(txtManuellerName, ">")
    If Pos <> 0 Then
        'MsgBox "Ein Dateiname darf keines der folgenden Zeichen enthalten: \/:*?""<>|"
        MsgBox LoadResString(1367 + Sprache)
        Exit Sub
    End If
    Pos = InStr(txtManuellerName, "|")
    If Pos <> 0 Then
        'MsgBox "Ein Dateiname darf keines der folgenden Zeichen enthalten: \/:*?""<>|"
        MsgBox LoadResString(1367 + Sprache)
        Exit Sub
    End If
    
    'temp = Dir(VordemJahr & JahrInFilename & "\" & txtManuellerName)
    If file_path_exist(VordemJahr & JahrInFilename & "\" & txtManuellerName) = True Then
    'If temp <> "" Then
        'Es ist immer noch ein Duplikat
        'MsgBox "Es ist immer noch ein Dateiname der zum Duplikat führen würde"
        MsgBox "Es ist immer noch ein Dateiname der zum Duplikat führen würde"
        Exit Sub
    Else
        Form1.gstrNeuerName = VordemJahr & JahrInFilename & "\" & txtManuellerName
        Unload Me
    End If
    gblnAbbrechen = False
End Sub

Private Sub Form_Load()
    Dim Fotodatei As String
    Dim start As Long
    
    Call AnpassenNutzerWunsch(Me)                                       'Gerbing 11.03.2017
    Label1.Caption = LoadResString(1301 + Sprache)    'Im Zielordner gibt es bereits eine Datei mit einem gleichen Namen
    Label2.Caption = LoadResString(1302 + Sprache)    'Ändern Sie jetzt den Dateinamen, so daß kein Duplikat entsteht.
    Label3.Caption = LoadResString(1303 + Sprache)    'Sie dürfen nur den Dateinamen-Anteil ändern, der nach dem Jahr steht
    btnOK.Caption = LoadResString(1304 + Sprache)     '&Geänderten Name benutzen
    btnAbbrechen.Caption = LoadResString(3013 + Sprache) '&Abbrechen
    Me.Caption = LoadResString(1305 + Sprache)         'Fehler Doppelter Name
    
    lblDuplikatName = Form1.gstrNeuerName
    
    Fotodatei = Form1.gstrNeuerName
    start = 1
    Do
        Pos = InStr(start, Fotodatei, "\") 'hintersten \ suchen
        If Pos = 0 Then Exit Do
        start = Pos + 1
    Loop
    JahrInFilename = Mid(Fotodatei, start - 5, 4)
    VordemJahr = Mid(Fotodatei, 1, start - 1 - 5)
    NachDemJahr = Mid(Fotodatei, start, Len(Fotodatei) - start + 1)
    
    txtManuellerName = NachDemJahr
End Sub
