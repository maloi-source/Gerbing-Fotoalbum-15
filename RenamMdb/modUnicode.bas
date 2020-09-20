Attribute VB_Name = "modUnicode"
Option Explicit

Public Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private opFile As OPENFILENAME

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameW" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameW" (ByRef pOpenfilename As OPENFILENAME) As Long

Private Declare Sub GetChar Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub PutChar Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal Destination As Long, ByRef Source As Long, ByVal Length As Long)


Public Property Get LCh(ByRef Text As String) As Long
    GetChar LCh, StrPtr(Text), 2
End Property

Public Property Let LCh(ByRef Text As String, ByVal NewValue As Long)
    PutChar StrPtr(Text), NewValue, 2
End Property

Public Property Get RCh(ByRef Text As String) As Long
    GetChar RCh, StrPtr(Text) + LenB(Text) - 2, 2
End Property

Public Property Let RCh(ByRef Text As String, ByVal NewValue As Long)
    PutChar StrPtr(Text) + LenB(Text) - 2, NewValue, 2
End Property

Public Property Get MCh(ByRef Text As String, ByVal Position As Long) As Long
    GetChar MCh, StrPtr(Text) + ((Position - 1) * 2), 2
End Property

Public Property Let MCh(ByRef Text As String, ByVal Position As Long, ByVal NewValue As Long)
    PutChar StrPtr(Text) + ((Position - 1) * 2), NewValue, 2
End Property

Public Function ShowOpenUnicode(Fm As Form) As String
    With opFile
        .flags = &H2 Or &H4
        .hInstance = App.hInstance
        .hWndOwner = Fm.hWnd
        .lpstrFilter = StrConv(("dia Files" & Chr(0) & "*.dia" & Chr(0) & Chr(0)), vbUnicode)
        '.lpstrFilter = StrConv(("All Files" & Chr(0) & "*.*" & Chr(0) & Chr(0)), vbUnicode)
        .lpstrTitle = StrConv("Open File", vbUnicode)
        .lpstrFile = StrConv(String(256, Chr(0)), vbUnicode)
        .nMaxFile = 512
        .lStructSize = Len(opFile)
    End With
    Call GetOpenFileName(opFile)
    
    ShowOpenUnicode = opFile.lpstrFile
End Function

Public Function ShowOpenUnicodeFotosMdb(Fm As Form) As String
    With opFile
        .flags = &H2 Or &H4
        .hInstance = App.hInstance
        .hWndOwner = Fm.hWnd
        .lpstrFilter = StrConv(("mdb Files" & Chr(0) & "*.mdb" & Chr(0) & Chr(0)), vbUnicode)
        '.lpstrFilter = StrConv(("All Files" & Chr(0) & "*.*" & Chr(0) & Chr(0)), vbUnicode)
        .lpstrTitle = StrConv("Fotos.mdb File", vbUnicode)
        .lpstrFile = StrConv(String(256, Chr(0)), vbUnicode)
        .nMaxFile = 512
        .lStructSize = Len(opFile)
    End With
    Call GetOpenFileName(opFile)
    
    ShowOpenUnicodeFotosMdb = opFile.lpstrFile
End Function

Public Function ConvertFileName(sToConvert) As String
    Dim bFileName() As Byte
    Dim lRet As Long
    Dim sBuf As String
    'Get rid of the trailing Null characters
    sToConvert = Left$(sToConvert, InStr(sToConvert, (Chr(0) & Chr(0))) - 1)
    
    If Len(sToConvert) Mod 2 <> 0 Then
        sToConvert = sToConvert & Chr(0) 'If the file has an ANSI extension or just an ANSI last character of
                                         'a file with no extension, add one on the end
                                         'If we don't add it, the string will end one character too short
    End If
    
    bFileName = StrConv(sToConvert, vbFromUnicode) 'Put the string into a byte array
    
    sBuf = ""
    For lRet = 0 To Len(sToConvert) - 1 Step 2
        'At this point, the unicode characters will show up in sBuf as ?, but, when we actually go
        'to use this in the FSO function, it will find the right file
        sBuf = sBuf & StrConv(Chr(bFileName(lRet)) & Chr(bFileName(lRet + 1)), vbFromUnicode)
    Next

    'And return the string for use
    ConvertFileName = sBuf
End Function

'Purpose:Returns True if string has a Unicode char.
Public Function IsStringUnicode(S As String) As Boolean
   Dim i As Long
   Dim bLen As Long
   Dim Map() As Byte

   If LenB(S) Then
        Map = S
        bLen = UBound(Map)
        For i = 1 To bLen Step 2
            If (Map(i) > 0) Then
                IsStringUnicode = True
                Exit Function
            End If
        Next
   End If
End Function

