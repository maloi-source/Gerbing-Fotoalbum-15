Attribute VB_Name = "bFolder"
Option Explicit
 
'ist unicode fähig
' == Dialog-Einstellungen ================================
 
' String, der vor dem aktuell ausgewählen Verzeichnis angezeigt wird,
' falls der ShowCurrentPath-Paramter True ist.
Private Const DIALOG_CURRENT_SELECTION_TEXT As String = "Auswahl: "
 
 
' == API-Deklarationen ===================================

Public Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatW" _
    (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderW" _
    (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListW" _
    (ByVal pidList As Long, ByVal lpBuffer As Long) As Long

 
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
 
Private Type Size
  cx As Long
  cy As Long
End Type
  
Private Declare Function SendMessage Lib "user32" _
  Alias "SendMessageW" ( _
  ByVal hWnd As Long, _
  ByVal wMsg As Long, _
  ByVal wParam As Long, _
  lParam As Any) As Long
 
Private Declare Sub CopyMemory Lib "kernel32" _
  Alias "RtlMoveMemory" ( _
  pDest As Any, _
  pSource As Any, _
  ByVal dwLength As Long)
 
Private Declare Function ILCreateFromPath Lib "shell32" _
  Alias "#157" ( _
  ByVal sPath As String) As Long
 
Private Declare Function LocalAlloc Lib "kernel32" ( _
  ByVal uFlags As Long, _
  ByVal uBytes As Long) As Long
 
Private Declare Function LocalFree Lib "kernel32" ( _
  ByVal hMem As Long) As Long
 
Private Declare Function lstrcpyA Lib "kernel32" ( _
  lpString1 As Any, _
  lpString2 As Any) As Long
 
Private Declare Function lstrlenA Lib "kernel32" ( _
  lpString As Any) As Long
 
Private Declare Function FindWindowEx Lib "user32.dll" _
  Alias "FindWindowExW" ( _
  ByVal hWnd1 As Long, _
  ByVal hWnd2 As Long, _
  ByVal lpsz1 As String, _
  ByVal lpsz2 As String) As Long
 
Private Declare Function GetWindowDC Lib "user32.dll" ( _
  ByVal hWnd As Long) As Long
 
Private Declare Function GetWindowRect Lib "user32.dll" ( _
  ByVal hWnd As Long, _
  ByRef lpRect As RECT) As Long
 
Private Declare Function GetTextExtentPoint Lib "gdi32.dll" _
  Alias "GetTextExtentPointW" ( _
  ByVal hdc As Long, _
  ByVal lpszString As String, _
  ByVal cbString As Long, _
  ByRef lpSize As Size) As Long
 
Private Declare Function PathCompactPath Lib "shlwapi.dll" _
  Alias "PathCompactPathW" ( _
  ByVal hdc As Long, _
  ByVal pszPath As String, _
  ByVal dx As Long) As Long
 
Private Const MAX_PATH = 260
 
Private Const WM_USER = &H400
 
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED As Long = 2
Private Const BFFM_SETSTATUSTEXTA As Long = (WM_USER + 100)
Private Const BFFM_SETSTATUSTEXTW As Long = (WM_USER + 104)
Private Const BFFM_ENABLEOK As Long = (WM_USER + 101)
Private Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Private Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)
 
Private Const BIF_NEWDIALOGSTYLE As Long = &H40
Private Const BIF_RETURNONLYFSDIRS As Long = &H1
Public Const BIF_BROWSEINCLUDEFILES As Long = &H4000
Private Const BIF_STATUSTEXT As Long = &H4
 
Private Const LMEM_FIXED = &H0
Private Const LMEM_ZEROINIT = &H40
Private Const lPtr = (LMEM_FIXED Or LMEM_ZEROINIT)

' Zeigt den BrowseForFolder-Dialog an.
Public Function BrowseForFolder(DialogText As String, _
  DefaultPath As String, _
  OwnerhWnd As Long, _
  Optional ShowCurrentPath As Boolean = True, _
  Optional RootPath As Variant, _
  Optional NewDialogStyle As Boolean = False, _
  Optional IncludeFiles As Boolean = False) As String
 
  ' Parameter:
  '    o DialogText        Dialogtext, der oben im Dialog angezeigt wird.
  '    o DefaultPath       Standardmäßig ausgewähltes Verzeichnis.
  '    o OwnerhWnd         hWnd des übergeordneten Fensters (in den meisten
  '                          Fällen Me.hWnd).
  '    o ShowCurrentPath   Legt fest, ob die aktuelle Verzeichnisauswahl
  '                          angezeigt werden soll. Verfügbar ab
  '                          Internet Explorer 4.0 (-> PathCompactPath).
  '    o RootPath          Root-Verzeichnis. Wird es angegeben, werden nur die
  '                          Ordner unterhalb dieses Verzeichnisses angezeigt.
  '    o NewDialogStyle    Legt fest, ob der Dialog in der neuen Darstellung
  '                          angezeigt werden soll (Dialog kann vergrößert/
  '                          verkleinert werden, es ist eine Schaltfläche zum
  '                          Anlegen eines neuen Ordners vorhanden, es können
  '                          Dateioperationen wie löschen etc. ausgeführt
  '                          werden, ...). Ist dieser Parameter True, hat der
  '                          Parameter ShowCurrentPath keine Wirkung. Verfügbar
  '                          unter WinME und Betriebsystemen ab Win2000.
  '    o IncludeFiles      Legt fest, ob auch Dateien im Dialog angezeigt und
  '                          ausgewählt werden können.
  '                        Verfügbar ab Win98 und Internet Explorer 4.0 (bei
  '                          frühreren Windowsversionen muss IE4 inkl. der
  '                          Integrated Shell installiert sein).
 
  Dim biBrowseInfo As BrowseInfo
  Dim lPIDL As Long
  Dim sBuffer As String
  Dim lBufferPointer As Long
 
  With biBrowseInfo
    ' Handle des übergeordneten Fensters
    .hWndOwner = OwnerhWnd
 
    ' PIDL des Rootordners zuweisen
    If Not IsMissing(RootPath) Then .pIDLRoot = PathToPIDL(RootPath)
 
    ' Dialogtext zuweisen
    If ShowCurrentPath And DialogText = "$" Then DialogText = "" ' Wird intern nicht zugelassen
    .lpszTitle = lstrcat(StrPtr(DialogText), StrPtr(""))
    '.lpszTitle = DialogText
 
    ' Stringbuffer für aktuell selektierten Pfad zuweisen
    If ShowCurrentPath Then .pszDisplayName = sBuffer
 
    ' Dialogeinstellungen zuweisen
    .ulFlags = BIF_RETURNONLYFSDIRS + _
      IIf(ShowCurrentPath, BIF_STATUSTEXT, 0) + _
      IIf(NewDialogStyle, BIF_NEWDIALOGSTYLE, 0) + _
      IIf(IncludeFiles, BIF_BROWSEINCLUDEFILES, 0)
 
    ' Callbackfunktion-Adresse zuweisen
    .lpfnCallback = FARPROC(AddressOf CallbackString)
 
    ' PIDL des vorselektierten Ordnerpfades zuweisen (wird im
    ' lpData-Parameter an die Callback-Funktion weitergeleitet)
    .lParam = PathToPIDL(DefaultPath)
  End With
 
  ' BrowseForFolder-Dialog anzeigen
  lPIDL = SHBrowseForFolder(biBrowseInfo)
 
  If lPIDL Then
    ' Stringspeicher reservieren
    sBuffer = Space$(MAX_PATH)
 
    ' Selektierten Pfad aus der zurückgegebenen PIDL ermitteln
    'SHGetPathFromIDList lPIDL, sBuffer
    SHGetPathFromIDList lPIDL, StrPtr(sBuffer)
 
    ' Nullterminierungszeichen des Strings entfernen
    sBuffer = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
 
    ' Selektierten Pfad zurückgeben
    BrowseForFolder = sBuffer
 
    ' Reservierten Task-Speicher wieder freigeben
    Call CoTaskMemFree(lPIDL)
  End If
 
  ' Stringspeicher wieder freigeben
  If ShowCurrentPath Then Call LocalFree(lBufferPointer)
End Function
Private Function CallbackString(ByVal hWnd As Long, ByVal uMsg As Long, _
  ByVal lParam As Long, ByVal lpData As Long) As Long
 
  ' Callback-Funktion des BrowseForFolder-Dialogs. Wird bei
  ' eintretenden Ereignissen des Dialogs aufgerufen.
 
  Dim sBuffer As String
  Dim lStaticWnd As Long
  Dim lStaticDC As Long
  Dim sPath As String
  Dim rctStatic As RECT
  Dim szTextSize As Size
 
  ' Meldungen herausfiltern
  Select Case uMsg
  Case BFFM_INITIALIZED
    ' Dialog wurde initialisiert
 
    ' Standardmäßig markierten Pfad (dessen PIDL wurde in lpData
    ' übergeben) im Dialog selektieren
    Call SendMessage(hWnd, BFFM_SETSELECTIONA, False, ByVal lpData)
  Case BFFM_SELCHANGED
    ' Selektion hat sich geändert
 
    ' Stringspeicher reservieren
    sBuffer = Space$(MAX_PATH)
 
    ' Aktuell selektierten Pfad ermitteln und anzeigen, wenn möglich
    If SHGetPathFromIDList(lParam, StrPtr(sBuffer)) Then
      ' Temporäre Zeichenfolge an das Anzeigelabel senden, um
      ' dessen Handle anhand dieser Zeichenfolge ermitteln zu können
      SendMessage hWnd, BFFM_SETSTATUSTEXTA, 0&, ByVal "$"
 
      ' Handle und DeviceContext des Anzeigelabels ermitteln
      lStaticWnd = FindWindowEx(hWnd, ByVal 0&, ByVal "Static", ByVal "$")
      lStaticDC = GetWindowDC(lStaticWnd)
 
      ' Abmessungen des Anzeigelabels ermitteln
      GetWindowRect lStaticWnd, rctStatic
 
      ' Textabmessungen der Zeichenfolge "Auswahl: " im Anzeigelabel
      ' ermitteln
      GetTextExtentPoint lStaticDC, ByVal DIALOG_CURRENT_SELECTION_TEXT, _
        ByVal Len(DIALOG_CURRENT_SELECTION_TEXT), szTextSize
 
      ' Anzuzeigenden Pfad auf die Abmessungen des Anzeigelabels
      ' kürzen; falls dies nicht möglich ist, gesamten Pfad anzeigen
      sPath = sBuffer
      If PathCompactPath(ByVal lStaticDC, sPath, ByVal (rctStatic.Right - _
        rctStatic.Left - szTextSize.cx + 80)) = 0 Then sPath = sBuffer
 
      ' Nullterminierung entfernen
      sPath = Left$(sPath, InStr(1, sPath, vbNullChar) - 1)
 
      ' Pfad im Dialog anzeigen
      Call SendMessage(hWnd, BFFM_SETSTATUSTEXTA, 0&, _
        ByVal DIALOG_CURRENT_SELECTION_TEXT & sPath)
    Else
      ' Pfadanzeige leeren
      SendMessage hWnd, BFFM_SETSTATUSTEXTA, 0&, ByVal ""
    End If
  End Select
End Function
Private Function FARPROC(FunctionPointer As Long) As Long
  ' Funktion wird benötigt, um Funktions-Adresse ermitteln
  ' zu können, dessen Adresse mit AddressOf übergeben und
  ' anschließend wieder zurückgegeben wird.
 
  FARPROC = FunctionPointer
End Function
' Gibt die lPIDL zum übergebenen Pfad zurück.
Private Function PathToPIDL(ByVal sPath As String) As Long
  Dim lRet As Long
 
  lRet = ILCreateFromPath(sPath)
  If lRet = 0 Then
    sPath = StrConv(sPath, VbStrConv.vbUnicode)
    lRet = ILCreateFromPath(sPath)
  End If
 
  PathToPIDL = lRet
End Function

