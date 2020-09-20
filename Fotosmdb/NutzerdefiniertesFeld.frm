VERSION 5.00
Begin VB.Form ND 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Nutzerdefiniertes Datenbank-Feld anlegen"
   ClientHeight    =   2484
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   11796
   Icon            =   "NutzerdefiniertesFeld.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2484
   ScaleWidth      =   11796
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   1092
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "NutzerdefiniertesFeld.frx":038A
      Top             =   120
      Width           =   7692
   End
   Begin VB.CommandButton btnAbbrechen 
      Caption         =   "&Abbrechen"
      Height          =   372
      Left            =   8040
      TabIndex        =   6
      Top             =   840
      Width           =   3492
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Height          =   372
      Left            =   8040
      TabIndex        =   5
      Top             =   120
      Width           =   3492
   End
   Begin VB.TextBox txtLänge 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1031
         SubFormatType   =   0
      EndProperty
      Height          =   372
      Left            =   6000
      TabIndex        =   4
      Top             =   1800
      Width           =   492
   End
   Begin VB.TextBox txtFeldname 
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   3492
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Feldlänge:"
      Height          =   372
      Left            =   6000
      TabIndex        =   3
      Top             =   1440
      Width           =   1692
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Feldtyp=Text"
      Height          =   372
      Left            =   3720
      TabIndex        =   2
      Top             =   1800
      Width           =   2052
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Feldname:"
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   1332
   End
End
Attribute VB_Name = "ND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAbbrechen_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    Dim SQL As String
    Dim Msg As String
    
    'Bei einer schreibgeschützten Datenbank kann man keine Felder erzeugen
    If txtFeldname = "" Then
        'MsgBox "Sie haben keinen Feldname eingegeben"
        MsgBox LoadResString(1369 + Sprache)
        Exit Sub
    End If
    If Not IsNumeric(txtLänge) Then                                 'Gerbing 07.05.2007
        'MsgBox "die Feldlänge muss zwischen 1 und 255 liegen"
        MsgBox LoadResString(1370 + Sprache)
        txtLänge.SetFocus
        Exit Sub
    End If
    If txtLänge > 255 Or txtLänge < 1 Then                          'Gerbing 07.05.2007
        'MsgBox "die Feldlänge muss zwischen 1 und 255 liegen"
        MsgBox LoadResString(1370 + Sprache)
        txtLänge.SetFocus
        Exit Sub
    End If

    If gblnSQLServerVersion = True Then
        'Beim sql server gilt andere Syntax
        SQL = "ALTER TABLE Fotos ADD " & txtFeldname & " nvarchar(" & txtLänge & ");"           'Gerbing 05.09.2016
        On Error Resume Next
        Form1.DBsql.Execute SQL
    Else
        'On Error Resume Next
        Form1.DBsql.Close
        On Error GoTo 0
        Form1.DBsql.mode = adModeReadWrite
        Form1.DBsql.Open Form1.DBsql.ConnectionString
        SQL = "ALTER TABLE Fotos ADD COLUMN [" & txtFeldname & "] TEXT(" & txtLänge & ");"
        On Error Resume Next
        Form1.DBsql.Execute SQL
    End If
    If ERR.Number <> 0 Then
        Msg = "Errornumber=" & ERR.Number & vbNewLine
        Msg = Msg & "Errortext=" & ERR.Description
        MsgBox Msg
    End If
        Msg = "Sie müssen jetzt das Programm neu starten"
        Msg = LoadResString(2207 + Sprache)
        MsgBox Msg
        End
    Unload Me
End Sub

Private Sub Form_Load()
    Call AnpassenNutzerWunsch(Me)                                       'Gerbing 11.03.2017
    Me.Caption = LoadResString(1345 + Sprache)  'Nutzerdefiniertes Datenbank-Feld anlegen
    If gblnSQLServerVersion = True Then
        Text1.Text = LoadResString(1837 + Sprache)      'Bitte beachten Sie, dass Sie die nutzerdefinierten Datenbank-Felder nicht mit Mitteln dieses Programms wieder entfernen können. Für solche Zwecke müssen Sie Microsoft SQL Server benutzen.
    Else
        Text1.Text = LoadResString(1346 + Sprache)      'Bitte beachten Sie, dass Sie die nutzerdefinierten Datenbank-Felder nicht mit Mitteln dieses Programms wieder entfernen können. Für solche Zwecke müssen Sie Microsoft Access benutzen.
    End If
    Label1.Caption = LoadResString(1347 + Sprache)      'Feldname:
    Label2.Caption = LoadResString(1348 + Sprache)      'Feldtyp=Text
    Label3.Caption = LoadResString(1349 + Sprache)      'Feldlänge:
    btnOK.Caption = LoadResString(3001 + Sprache)       '&OK
    btnAbbrechen.Caption = LoadResString(3013 + Sprache)        '&Abbrechen
End Sub
