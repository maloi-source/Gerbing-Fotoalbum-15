Attribute VB_Name = "Module2"
Option Explicit
' ###########################################
'
'                      MODUL "CommonDialog"
'
'                     (c) Ingo Steinhaus 2000
'                     ingo.steinhaus@gmx.de
'
'  Funktionen zur Anzeige der Windows-Standarddialoge "Öffnen"
'  und "Speichern"
'
'  Dieses Modul ist urheberrechtlich geschützte Freeware.
'  Die originale Copyright-Meldung darf nicht entfernt oder ver-
'  ändert werden. Der Quelltext darf nicht verändert werden-
'
'############################################

Public Const MAX_PATH = 260

'**********************************************************
Rem Die Datenstruktur "OpenFilename" dient der Konfiguration
Rem des Dialogs.

Private Type OpenFilename
    lStructSize As Long
    'Größe der Datenstruktur. Kann mit Len() bestimmt werden.
    hWndOwner As Long
    'Handle des Besitzers (mit GetActiveWindow() abfragen).
    hInstance As Long
    'Handle der Dialogfeldvorlage, wenn OFN_ENABLETEMPLATEHANDLE
    'in Flags gesetzt ist. Wenn OFN_EXPLORER gesetzt ist, wird der
    'Dialog vom Standardialog des Explorers abgeleitet. Andernfalls
    'wird ein Dialog im Windows-3.x-Stil erzeugt.
    lpstrFilter As String
    'Ein VB-String mit paarweise angeordneten nullterminierten Strings.
    'Der letzte nullterminierte String muß mitr einem weiteren NULL-Zeichen
    'abgeschlossen werden.
    'Ein Filter besteht aus zwei nullterminierten Strings. Der erste enthält
    'die Zeichenkette, die im Kombifeld "DateiTyp" angezeigt wird, der zweite
    'die zugehörtigen Dateimasken wie z. B. "*.doc".
    'Beispiel: "Word-Dokumente" + Chr$(0) + "*.doc" + Chr$(0)
    'Sie können mehrere Dateimasken durch Semikola abtrennen.
    'Beispiel: "Grafiken" + Chr$(0) + "*.bmp;*.jpg;*.gif" + Chr$(0)
    lpstrCustomFilter As String
    'Ein VB-String mit dem im Kombifeld "DateiTyp" ausgewählten Filter.
    nMaxCustFilter As Long
    'Die Größe von lpstrCustomFilter.
    nFilterIndex As Long
    'Der 1-basierte Index des im Kombifeld "DateiTyp" ausgewählten Filters.
    lpstrFile As String
    'Ein VB-String mit dem ausgewählten Dateinamen inkl. Laufwerk und Pfad.
    'Der String muß vorher in der entsprechenden Größe erzeugt werden.
    'Er kann vor dem Aufruf des Dialogs mit dem Namen einer existierenden
    'Datei belegt werden.
    nMaxFile As Long
    'Die Größe von lpstrFile.
    lpstrFileTitle As String
    'Ein VB-String mit dem ausgewählten Dateinamen ohne Laufwerk und Pfad.
    nMaxFileTitle As Long
    'Die Größe von lpstrFileTitle.
    lpstrInitialDir As String
    'Ein VB-String mit dem Pfadnamen des Ordners, dessen Inhalt der Dialog
    'beim Anzeigen darstellen soll.
    lpstrTitle As String
    'Ein VB-String mit Titel des Dialogfeldes.
    Flags As Long
    'Flags, die die Anzeigeoptionen des Dialogfeldes bestimmen
    nFileOffset As Integer
    'Index zum Beginn des ersten Dateinamens in lpstrFile.
    nFileExtension As Integer
    'Index zum Beginn der Dateierweiterung in lpstrFile.
    lpstrDefExt As String
    'Die Standarderweiterung, die an einen Dateinamen vergeben wird, wenn
    'er keine Erweiterung besitzt.
    lCustData As Long
    'Ein Zeiger auf anwendungsspezifiasche Daten, für die Rückruffunktion.
    lpfnHook As Long
    'Adresse einer Rückruffunktion, die in der Anwendung definiert wird.
    'Sie können hier NULL eintragen.
    lpTemplateName As String
    'Der Name der Dialogfeldvorlage (siehe hInstance)
End Type

'**********************************************************
Rem Die folgenden Konstanten sind die erlaubten Werte für
Rem OpenFilename->Flags.

Public Const OFN_ALLOWMULTISELECT = &H200
'Zeigt ein Dialogfeld mit der Möglichkeit, mehrere Dateien auszuwählen.
'In diesem Fall enthält lpstrFile den Pfad und anschließend alle Dateinamen.
'nFileOffset zeigt auf den Index des ersten Dateinamens nach der Pfadangabe.
'lpstrFile enthält alle Dateinamen durch Chr$(0) getrennt. Em Ende folgt ein
'zweites Chr$(0). Bei alten Win-3.x-Dialoge) sind die Dateinamen durch
'Leerzeichen getrennt. Diese Variante kennt keine langen Dateinamen.

Public Const OFN_CREATEPROMPT = &H2000
'Zeigt eine Meldung, wenn die Datei nicht existiert und fragt den Anwender, ob
'sie erzeugt werden soll.

Public Const OFN_ENABLEHOOK = &H20
'Aktiviert die Rückruffunktion lpfnHook.

Public Const OFN_ENABLETEMPLATE = &H40
'Aktiviert die Dialogfeldvorlage.

Public Const OFN_ENABLETEMPLATEHANDLE = &H80
'Aktiviert die Dialogfeldvorlage.

Public Const OFN_EXPLORER = &H80000
'Nutzt Explorer-Dialoge. Diese Einstellung ist die Vorgabe, selbst wenn Sie
'dieses Flag nicht angeben. Für alte Win-3.x-Dialoge müssen Sie das Flag
'löschen.
'Sie müssen es in den folgenden Fällen setzen:
'- bei OFN_ALLOWMULTISELECT.
'- wenn Sie Dialogfeldvorlagen und Rückruffunktionen benutzen.

Public Const OFN_EXTENSIONDIFFERENT = &H400&
'Gibt an, dass der Anwender einen Dateinamen mit einer anderen Erweiterung
'als lpstrDefExt eingeben kann.

Public Const OFN_FILEMUSTEXIST = &H1000
'Gibt an, dass der Anwender nur die Namen von existierenden Dateien eingeben
'kann. Andernfalls wird eine Warnmeldung ausgegeben.
'OFN_PATHMUSTEXIST muß ebenfalls gesetzt werden.

Public Const OFN_HIDEREADONLY = &H4&
'Versteckt das Kontrollkästchen "Nur lesen".

Public Const OFN_LONGNAMES = &H200000
'Aktiviert die Unterstützung von langen Dateinamen in den alten Win-3.x-Dialogen.

Public Const OFN_NOCHANGEDIR = &H8&
'Stellt das ursprüngliche Verzeichnis bei Ende des Dialoges wieder her, wenn
'der Anwender anderes Verzeichnis eingestellt hat.

Public Const OFN_NODEREFERENCELINKS = &H100000
'Weist das Dialogfeld an, bei einer markierten Verknüpfung Namen und Pfad der
'Verknüpungsdatei zurückzugeben, snstatt Namen und Pfad der Datei, auf die die
'Verknüpfung verweist.

Public Const OFN_NOLONGNAMES = &H40000
'Deaktiviert die Unterstützung von langen Dateinamen in den alten
'Win-3.x-Dialogen.

Public Const OFN_NONETWORKBUTTON = &H20000
'Versteckt die Schaltfläche "Netzwerk".

Public Const OFN_NOTESTFILECREATE = &H10000
'Gibt an, dass keine Testdatei erzeugt wird, bevor der Dialog endet. In diesem
'Fall überprüft das Dialogfeld nicht auf Schreibschutz, Platzmangel auf dem
'Datenträger oder korrekten Netzwerkzugriff.

Public Const OFN_OVERWRITEPROMPT = &H2&
'Gibt im Dialog "Speichern" eine Warnmeldung aus, wenn die Datei bereits
'existiert und durch das Speichern überschrieben wird.

Public Const OFN_PATHMUSTEXIST = &H800
'Gibt an, dass der Anwender nur die Namen von existierenden Verzeichnissen
'eingeben kann. Andernfalls wird eine Warnmeldung ausgegeben.

Public Const OFN_READONLY = &H1
'Gibt an, das das Kontrollkästchen "Nur Lesen" angekreuzt ist, wenn der Dialog
'angezeigt wird.

Public Const OFN_SHAREAWARE = &H4000
'Gibt an, dass die Funktion fehlschlägt, wenn ein Netzwerkfehler auftritt.

Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHAREWARN = 0
Public Const OFN_SHARENOWARN = 1

Public Const OFN_SHOWHELP = &H10
'Zeigt im Dialogfeld den Hilfe-Schalter an. hwndOwner muß auf ein Fenster zeigen,
'das die Hilfe anzeigen kann. ExplorerDialoge senden die Nachricht CDN_HELP
'an die Rückruffunktion.

Public Const OFS_MAXPATHNAME = 128

'**********************************************************
Rem *** GetSaveFileName ***
Rem Funktion zum Anzeigen des Dialogs "Speichern"

Private Declare Function GetSaveFileName Lib "comdlg32" Alias _
    "GetSaveFileNameA" (lpOpenfilename As OpenFilename) As Long

'**********************************************************
Rem *** GetOpenFileName ***
Rem Funktion zum Anzeigen des Dialogs "Öffnen"

Private Declare Function GetOpenFileName Lib "comdlg32" Alias _
    "GetOpenFileNameA" _
    (lpOpenfilename As OpenFilename) As Long

'**********************************************************
Rem *** CommDlgExtendedError ***
Rem Funktion zum Ermitteln der Fehlernummer

Private Declare Function CommDlgExtendedError Lib "comdlg32" () As Integer

'**********************************************************
Rem *** GetActiveWindow ***
Rem Eine Funktion zum Ermitteln des Fenster-Handles.

Private Declare Function GetActiveWindow Lib "user32" () As Long

'**********************************************************
Rem *** PrepareFilter ***
Rem Eine Funktion zum Aufbereiten des Filters
Rem Beispielfilter: "Word Dokument (.doc)|*.doc|Word Vorlage (.dot)|*.dot|"
Rem Die Funktion ersetzt "|" durch Chr$(0) und fügt das abschließende
Rem Chr$(0) ein.

Public Function PrepareFilter(Flt$) As String
Const O$ = "|"
Dim Temp$
Dim i As Integer
    'Mit einer Kopie arbeiten
    Temp$ = Flt$
    'Beim ersten Zeichen beginnen
    i = 1
    '"|" gefunden?
    Do While InStr(i, Flt$, O$) <> 0
        'Alles bis zum ersten "|" kopieren und Chr$(0) anhängen
        PrepareFilter = PrepareFilter + _
                Mid(Temp$, i, InStr(i, Temp$, O$) - i) + vbNullChar
        'Index auf Zeichen nach "|" setzen
        i = InStr(i, Temp$, O$) + Len(O$)
    Loop
    'Evtl. Rest vom String und abschließendes Chr$(0) anhängen
    PrepareFilter = PrepareFilter + Right(Temp$, Len(Temp$) - i + 1) + vbNullChar
End Function

'**********************************************************
Rem *** GetSaveName ***
Rem Eine VB/VBA-Funktion als einfach zu nutzender Mantel für den
Rem Aufruf des Dialogs "Speichern".

Public Function GetSaveName(ByVal Filter$, ByVal DefExt$, ByVal InitialDir$) As String
Dim OFN As OpenFilename
Dim Temp$
Dim n As Integer

    'Bestimmen der Optionen für den Dialog
    With OFN
        'Größe der Struktur festlegen
        .lStructSize = Len(OFN)
        'Das aktive Fenster (= Word) wird zum Besitzer des Dialogs
        .hWndOwner = GetActiveWindow()
        'Der Filtzer wird vorbereitet
        .lpstrFilter = PrepareFilter(Filter$)
        'Speicher reservieren für kompletten Pfad
        .lpstrFile = String$(700, vbNullChar)
        'Größe des reservierten Speichers angeben
        .nMaxFile = 700
        'Speicher reservieren für Dateinamen
        .lpstrFileTitle = String$(MAX_PATH, vbNullChar)
        'Größe des reservierten Speichers angeben
        .nMaxFileTitle = MAX_PATH
        'Das Vorgabeverzeichnis bestimmen
        .lpstrInitialDir = InitialDir$
        'Der Titel des Dialoges
        .lpstrTitle = "Speichern"
        'Optionen bestimmen
        .Flags = OFN_EXTENSIONDIFFERENT Or _
            OFN_NOCHANGEDIR Or _
            OFN_OVERWRITEPROMPT _
            Or OFN_HIDEREADONLY
        'Standarderweiterung für die Dateien bestimmen
        .lpstrDefExt = DefExt$
    End With

    If GetSaveFileName(OFN) Then
        Temp$ = OFN.lpstrFile
        'Alles nach dem NULL-Zeichen verwerfen
        n = InStr(Temp$, vbNullChar)
        If n > 1 Then
            GetSaveName = left$(Temp$, n - 1)
        Else
            GetSaveName = ""
        End If
    Else
        GetSaveName = ""
    End If
End Function

'**********************************************************
Rem *** GetOpenName ***
Rem Eine VB/VBA-Funktion als einfach zu nutzender Mantel für den
Rem Aufruf des Dialogs "Öffnen".

Public Function GetOpenName(ByVal Filter$, ByVal InitialDir$, ByVal Title$) As String
Dim OFN As OpenFilename
Dim Temp$
Dim n As Integer

    'Bestimmen der Optionen für den Dialog
    With OFN
        .lStructSize = Len(OFN)
        .hWndOwner = GetActiveWindow()
        .lpstrFilter = PrepareFilter(Filter$)
        'Speicher reservieren
        .lpstrFile = String$(700, vbNullChar)
        .nMaxFile = 700
        .lpstrFileTitle = String$(MAX_PATH, vbNullChar)
        .nMaxFileTitle = MAX_PATH
        .lpstrInitialDir = InitialDir$
        '.lpstrTitle = "Öffnen"
        .lpstrTitle = Title
        .Flags = OFN_EXTENSIONDIFFERENT Or _
            OFN_NOCHANGEDIR Or _
            OFN_OVERWRITEPROMPT _
            Or OFN_HIDEREADONLY
    End With

    If GetOpenFileName(OFN) Then
        Temp$ = OFN.lpstrFile
        'Alles nach dem NULL-Zeichen verwerfen
        n = InStr(Temp$, vbNullChar)
        If n > 1 Then
            GetOpenName = left$(Temp$, n - 1)
        Else
            GetOpenName = ""
        End If
    Else
        GetOpenName = ""
    End If
End Function



