VERSION 5.00
Begin VB.Form frmBenutzeGespeicherteAbfrage 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Benutze Gespeicherte Abfrage - use saved query"
   ClientHeight    =   6252
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5292
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6252
   ScaleWidth      =   5292
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   5760
      Width           =   2055
   End
   Begin VB.TextBox txtSQLGespeicherteAbfrage 
      Height          =   2535
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   1
      Top             =   3120
      Width           =   5055
   End
   Begin VB.ListBox ListGespeicherteAbfragen 
      Height          =   2736
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmBenutzeGespeicherteAbfrage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim DBado As ADODB.Connection                                       'Gerbing 23.11.2017

Private Sub btnOK_Click()
    If txtSQLGespeicherteAbfrage = "" Then
        MsgBox "Sie haben keine Abfrage ausgewählt"
        Exit Sub
    End If
    Renam.BGASQL = txtSQLGespeicherteAbfrage
    Unload Me
End Sub

Private Sub Form_Load()
    Dim intLoop As Integer
    Dim pos As Long
    Dim Pos1 As Long
    Dim msg As String

    Call AnpassenNutzerWunsch(Me)                                       'Gerbing 11.03.2017
    
    Set DBado = CreateObject("ADODB.Connection")                        'Gerbing 23.11.2017
    DBado.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AppPath & "\fotos.mdb"
    DBado.mode = adModeReadWrite
    DBado.Open DBado.ConnectionString
    
    txtSQLGespeicherteAbfrage.Locked = True
    For intLoop = 0 To DBado.QueryDefs.Count - 1
        ListGespeicherteAbfragen.AddItem DBado.QueryDefs(intLoop).Name
    Next intLoop
End Sub

Private Sub ListGespeicherteAbfragen_Click()
    Dim pos As Long
    Dim Pos1 As Long
    Dim msg As String
    
'   Fehlerkorrektur bei gespeicherten Abfragen:                         'Gerbing 19.01.2007
'   Bisher kam es zum Fehler, wenn im SQL-Text zb 'Fotos.BreitePixel' formuliert wurde.
'   Richtig muss formuliert werden 'BreitePixel'
'   Das Programm ersetzt 'Fotos.' durch ""
    txtSQLGespeicherteAbfrage = DBado.QueryDefs(ListGespeicherteAbfragen.ListIndex).SQL
    pos = InStr(1, txtSQLGespeicherteAbfrage, "Fotos" & ".", vbTextCompare)
    If pos <> 0 Then
'        msg = "Sie benutzen im SQL-String Formulierungen der Art Tabellenname.Feldname." & vbNewLine
'        msg = msg & "Solche Formulierungen würden zu Laufzeitfehlern führen." & vbNewLine
'        msg = msg & "Das Programm wird diese Formulierungen jetzt entfernen."
        msg = LoadResString(2315 + Sprache) & vbNewLine
        msg = msg & LoadResString(2316 + Sprache) & vbNewLine
        msg = msg & LoadResString(2317 + Sprache)
        MsgBox msg
        txtSQLGespeicherteAbfrage = Replace(txtSQLGespeicherteAbfrage, "Fotos" & ".", "", , , vbTextCompare)
    End If
    'Wenn im SQL-Text Select *, Feldname FROM ..... formuliert wird, werden die bezeichneten Feldnamen
    'als darzustellendes Feld betrachtet und stehen im Recordset vor den üblichen Standardfeldern.
    'Das bringt die gespeicherten Feldbreiten durcheinander
    'Der nutzer erhält eine Warnung
    pos = InStr(1, txtSQLGespeicherteAbfrage, "Select *", vbTextCompare)
    Pos1 = InStr(1, txtSQLGespeicherteAbfrage, "FROM", vbTextCompare)
    If pos <> 0 And Pos1 <> 0 Then
        If Pos1 - pos > 10 Then
'            msg = "Vermeiden Sie im SQL-String Formulierungen wie Select *, Feldname FROM..." & vbNewLine
'            msg = msg & "Dadurch würden die gespeicherten Feldbreiten auf falsche Felder angewendet." & vbNewLine
'            msg = msg & "Formulieren Sie besser Select * FROM..."
'            msg = msg & "Das Programm wird diese Formulierungen jetzt entfernen."
            msg = LoadResString(2318 + Sprache) & vbNewLine
            msg = msg & LoadResString(2319 + Sprache) & vbNewLine
            msg = msg & LoadResString(2320 + Sprache) & vbNewLine                       'Gerbing 07.11.2007
            msg = msg & LoadResString(2317 + Sprache)                                   'Gerbing 07.11.2007
            MsgBox msg
            txtSQLGespeicherteAbfrage = "Select * " & Right(txtSQLGespeicherteAbfrage, Len(txtSQLGespeicherteAbfrage) - Pos1 + 1)   'Gerbing 07.11.2007
        End If
    End If
'   Warnung bei gespeicherten Abfragen:                                                 'Gerbing 07.11.2007
'   zwischen 'Select... und ...FROM' muss ein '*' stehen, sonst kann man den String 'Select * FROM...'
'   nicht herstellen. Es kommt ein Warnhinweis
    If pos = 0 Then
'        msg = "Die Formulierung 'Select * FROM' wurde nicht gefunden." & vbNewLine
'        msg = msg & "Möglicherweise könnten Sie unerwartete Ergebnisse erhalten."
        msg = LoadResString(2328 + Sprache) & vbNewLine
        msg = msg & LoadResString(2329 + Sprache)
        MsgBox msg
    End If
End Sub
