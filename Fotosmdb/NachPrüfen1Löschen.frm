VERSION 5.00
Object = "{FCCB83BF-E483-4317-9FF2-A460758238B5}#1.2#0"; "CBLCtlsU.ocx"
Begin VB.Form NachPr�fen1L�schen 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Pr�fen1 - Die gefunden Datens�tze sind �berfl�ssig -> l�schen aus der Datenbank"
   ClientHeight    =   8028
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   8808
   Icon            =   "NachPr�fen1L�schen.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8028
   ScaleWidth      =   8808
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton btnAbbrechen 
      Caption         =   "&Abbrechen"
      Height          =   372
      Left            =   4680
      TabIndex        =   1
      Top             =   7440
      Width           =   4092
   End
   Begin VB.CommandButton btnL�schen 
      Caption         =   "markierten Datensatz &l�schen"
      Height          =   372
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Zum Markieren k�nnen Sie die Tasten Umsch und Strg zu Hilfe nehmen"
      Top             =   7440
      Width           =   4092
   End
   Begin CBLCtlsLibUCtl.ListBox lstZus�tzlicheDateien 
      Height          =   7212
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   8532
      _cx             =   15049
      _cy             =   12721
      AllowDragDrop   =   0   'False
      AllowItemSelection=   -1  'True
      AlwaysShowVerticalScrollBar=   -1  'True
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   0
      ColumnWidth     =   -1
      DisabledEvents  =   1048800
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
Attribute VB_Name = "NachPr�fen1L�schen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim NL As String
    Dim Msg As String

Private Sub btnAbbrechen_Click()
    Me.Hide
End Sub

Private Sub btnL�schen_Click()
    Dim i As Long
    Dim n As Long
    Dim strTemp As String
    Dim SQL As String
    Dim MyAppPath As String
    Dim NewFileName As String                                                           'Gerbing 10.11.2016
    Dim start As Long
    Dim Pos As Long
    Dim rc As Long
    
    If gblnSQLServerVersion = True Then
        MyAppPath = PublicLocationFotos
    Else
        MyAppPath = AppPath
    End If

    If gblnSchreibgesch�tzt = True Then                             'Gerbing 23.01.2007
        'msg = "Bei einer schreibgesch�tzten Datenbank ist diese Funktion nicht m�glich"
        Msg = LoadResString(2421 + Sprache)
        MsgBox Msg
        Exit Sub
    End If
    '1.Die Tabelle Temp_Haken f�llen mit allen Dateinamen aus lstZus�tzlicheDateien.Selected(i)
    '2.Aus Tabelle Fotos alle Dateinamen l�schen, die auch in Temp_Haken stehen
        
    '1.1.Die Tabelle Temp-Haken leer machen
    '1.2.Ein Recorset mit der leeren Tabelle Temp_Haken aufmachen
    If gblnSQLServerVersion = True Then
        'Zuerst aus der Tabelle Temp_Haken alle S�tze l�schen           'Gerbing 29.12.2011
        'beim SQL Server muss es hei�en 'Delete from table
        SQL = "DELETE From Temp_Haken"
        'SQL = "DELETE FROM " & LoadResString(2523 + Sprache)
    Else
        'Zuerst aus der Tabelle Temp_Haken alle S�tze l�schen           'Gerbing 30.09.2004
        SQL = "DELETE " & "Temp_Haken.* "
        SQL = SQL & " FROM " & "Temp_Haken;"
        'SQL = "DELETE " & LoadResString(2523 + Sprache) & ".* "
        'SQL = SQL & " FROM " & LoadResString(2523 + Sprache)
    End If
    Form1.DBsql.Execute SQL
    'dann leeres Recordset rstTempHaken �ffnen
    SQL = " SELECT " & "Temp_Haken.*"
    SQL = SQL & " FROM " & "Temp_Haken;"
    'SQL = " SELECT " & LoadResString(2523 + Sprache) & ".*"
    'SQL = SQL & " FROM " & LoadResString(2523 + Sprache)
    Set Form1.rstTempHaken = New ADODB.Recordset
    With Form1.rstTempHaken
        .ActiveConnection = Form1.DBsql
        .CursorType = adOpenDynamic
        .CursorLocation = adUseClient
        .Source = SQL
        .LockType = adLockOptimistic
        .Open
    End With
    '1.3.Die Tabelle Temp_Haken f�llen mit allen Dateinamen aus lstZus�tzlicheDateien.Selected(i)
    i = 0
    Screen.MousePointer = vbHourglass
    lstZus�tzlicheDateien.Visible = False                   'Gerbing 26.01.2006
    Do While lstZus�tzlicheDateien.ListItems.Count <> 0
        If lstZus�tzlicheDateien.ListItems(i).Selected Then
            strTemp = Replace(lstZus�tzlicheDateien.ListItems(i), MyAppPath, "+:")
            Form1.rstTempHaken.AddNew
            Form1.rstTempHaken.Fields("Merker") = 0         'Gerbing 03.03.2012
            Form1.rstTempHaken.Fields("Dateiname") = strTemp
'--------------------------------------------------------------------------------------------------------------------
'           Seit Version 14.2.2 speichere ich Thumbnails im Ordner ...\GerbingThumbs\...    'Gerbing 10.11.2016 Start
'           Beim L�schen von Datens�tzen, deren Foto nicht gefunden wurde, l�sche ich auch zugeh�rige Thumbnails
            'Wenn es Thumbnails im Ordner ...\GerbingThumbs\... gibt, l�sche ich diese
            NewFileName = lstZus�tzlicheDateien.ListItems(i)
            start = 1
            Do
                Pos = InStr(start, NewFileName, "\")
                If Pos = 0 Then Exit Do
                start = Pos + 1
            Loop
            NewFileName = left(NewFileName, start - 1) & "GerbingThumbs\" & Right(NewFileName, Len(NewFileName) - start + 1) & ".jpg" 'Gerbing 08.12.2016
            If file_path_exist(NewFileName) = True Then
                rc = file_delete(NewFileName, False, True) 'ohne Papierkorb, silent
            End If                                                                          'Gerbing 10.11.2016 End
'--------------------------------------------------------------------------------------------------------------------
            Form1.rstTempHaken.Update
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
    '2.Aus Tabelle Fotos alle Dateinamen l�schen, die auch in Temp_Haken stehen
    'Inkonsistenzabfrage                                                'Gerbing 30.09.2004
    'Die Inkonsistenzabfrage findet alle Dateinamen in Tabelle Fotos, die auch in Tabelle Temp_Haken eingetragen sind
    'SQL = "SELECT Fotos.Dateiname FROM Fotos LEFT JOIN Temp-Haken ON Temp_Haken.Dateiname = Fotos.Dateiname"
    'SQL = SQL & " WHERE (((Temp_Haken.Dateiname) not Is Null));"
    SQL = "SELECT Fotos." & LoadResString(1028 + Sprache) & " FROM Fotos LEFT JOIN Temp_Haken ON Temp_Haken.Dateiname = Fotos." & LoadResString(1028 + Sprache)
    SQL = SQL & " WHERE (((Temp_Haken.Dateiname) Is not Null));"
    On Error Resume Next
    Form1.rstsql.Close
    On Error GoTo 0
    With Form1.rstsql
        .Source = SQL
        .ActiveConnection = Form1.DBsql
        '.CursorType = adOpenStatic
        .CursorType = adOpenForwardOnly
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    If Not Form1.rstsql.EOF Then
        Do Until Form1.rstsql.EOF
            Form1.rstsql.Delete
            Form1.rstsql.Movenext
        Loop
    End If
    lstZus�tzlicheDateien.Visible = True                'Gerbing 26.01.2006
    Screen.MousePointer = vbDefault
    Form1.FehlerGefunden = False
    Form1.txtFehlerU.Text = ""
End Sub

Private Sub Form_Load()
    Call AnpassenNutzerWunsch(Me)                                       'Gerbing 11.03.2017
    Me.Caption = LoadResString(1462 + Sprache)  'Pr�fen1 - Die gefunden Datens�tze sind �berfl�ssig -> l�schen aus der Datenbank
    btnL�schen.Caption = LoadResString(1463 + Sprache)      'markierten Datensatz &l�schen
    btnAbbrechen.Caption = LoadResString(3013 + Sprache)            '&Abbrechen
    btnL�schen.ToolTipText = LoadResString(1432 + Sprache)           'Zum Markieren k�nnen Sie die Tasten Umsch und Strg zu Hilfe nehmen
    
    'lstZus�tzlicheDateien.MultiSelect = 2 muss in der Entwicklungsumgebung eingestellt werden
    If lstZus�tzlicheDateien.ListItems.Count <> 0 Then
        'lstZus�tzlicheDateien.ListIndex = 0
    End If
    NL = vbNewLine
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lstZus�tzlicheDateien.width = NachPr�fen1L�schen.width - 200
    lstZus�tzlicheDateien.height = NachPr�fen1L�schen.height - 1140
    btnL�schen.top = Me.height - 975
    btnAbbrechen.top = Me.height - 975
    On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    Me.Hide
End Sub

