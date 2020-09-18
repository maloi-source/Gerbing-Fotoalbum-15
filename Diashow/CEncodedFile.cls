VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEncodedFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    'Geschrieben von Philipp Stephani, 2003
    'Diese Klasse (CEncodedFile) kann kodierte Dateien laden und speichern und deren Inhalt als VB-Zeichenfolgen darstellen
    'Unterstützte Kodierungsformen: ANSI, UTF-32, UTF-16, UTF-8
    'Es werden jeweils BE und LE unterstützt. Wegen der natürlichen Beschränkung der VB-Zeichenfolgen können nur Zeichen des BMP gelesen und geschrieben werden.
    'Das Kodierungsschema wird an Hand der Signatur (BOM) am Anfang der Datei erkannt.
    'Die Klasse ist natürlich beliebig erweiterbar.
    
    Private Declare Sub ByteToPointer Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal Destination As Long, ByRef Source As Byte, ByVal Length As Long)
    Private Declare Sub PointerToByte Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Byte, ByVal Source As Long, ByVal Length As Long)
    Private Declare Sub ByteToLong Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Long, ByRef Source As Byte, ByVal Length As Long)
    Private Declare Sub LongToByte Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Byte, ByRef Source As Long, ByVal Length As Long)
    Private Declare Function DecodeString Lib "kernel32.dll" Alias "MultiByteToWideChar" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Byte, ByVal cbMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
    Private Declare Function EncodeStringPointer Lib "kernel32.dll" Alias "WideCharToMultiByte" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
    Private Declare Function EncodeStringByte Lib "kernel32.dll" Alias "WideCharToMultiByte" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Byte, ByVal cbMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
    
    Private Const CP_ACP As Long = 0
    Private Const CP_UTF8 As Long = 65001
    
    Public Enum EEncodingScheme
        EEncodingANSI
        EEncodingUTF8
        EEncodingUTF16BE
        EEncodingUTF16LE
        EEncodingUTF32BE
        EEncodingUTF32LE
    End Enum
    
    Private Enum EEncodingSignature
        ESignatureUTF8Mask = &HFFFFFF
        ESignatureUTF8 = &HBFBBEF
        ESignatureUTF16Mask = &HFFFF&
        ESignatureUTF16BE = &HFFFE&
        ESignatureUTF16LE = &HFEFF&
        ESignatureUTF32BE = &HFFFE0000
        ESignatureUTF32LE = &HFEFF&
    End Enum
    
    Private mPath As String
    Private mText As String
    Private mEncoding As EEncodingScheme
    
    Private Declare Function GetShortPathName Lib "kernel32" _
        Alias "GetShortPathNameW" ( _
        ByVal lpszLongPath As Long, _
        ByVal lpszShortPath As Long, _
        ByVal cchBuffer As Long) As Long


Public Sub Load(ByRef bUni() As Byte)
    Dim Handle As Long
    Dim Size As Long
    Dim buffer() As Byte
    Dim Signature As EEncodingSignature
    Dim Length As Long
    Dim Position As Long
    Dim i As Long
    
    'Handle = FreeFile
    'Open Path For Binary Access Read As #Handle
    Size = UBound(bUni) + 1
    ReDim buffer(0 To Size \ 2)
    For i = 0 To Size - 1
        buffer(i \ 2) = bUni(i)
        i = i + 1
    Next i
    'ReDim Buffer(0 To Size - 1)
    'Get #Handle, , Buffer
    'Close #Handle
    'mPath = Path
    ByteToLong Signature, buffer(0), 4
    If Signature = ESignatureUTF32BE Then
        mEncoding = EEncodingUTF32BE
        mText = String$((Size - 4) \ 4, 0)
        For Position = 4 To Size - 3 Step 4
            MCh(mText, Position \ 4) = (buffer(Position + 2) * &H100&) Or buffer(Position + 3)
        Next Position
    ElseIf Signature = ESignatureUTF32LE Then
        mEncoding = EEncodingUTF32LE
        mText = String$((Size - 4) \ 4, 0)
        For Position = 4 To Size - 1 Step 4
            MCh(mText, Position \ 4) = buffer(Position) Or (buffer(Position + 1) * &H100&)
        Next Position
    ElseIf (Signature And ESignatureUTF8Mask) = ESignatureUTF8 Then
        mEncoding = EEncodingUTF8
        Length = DecodeString(CP_UTF8, 0, buffer(3), Size - 3, 0, 0)
        mText = String$(Length, 0)
        DecodeString CP_UTF8, 0, buffer(3), Size - 3, StrPtr(mText), Length
        'Unter Windows 95 funktioniert "MultiByteToWideChar" nicht mit UTF-8-Eingaben.
        'Stattdessen kann die gleichnamige Funktion aus der Bibliothek "unicows.dll" (Microsoft Layer for Unicode) benutzt werden.
        'Die Deklaration und der Aufruf der beiden Funktionen ist bis auf die Bibliotheksangabe völlig gleich.
        'Will man die Bibliothek "unicows.dll" nicht mit ausliefern, kann folgender Algorithmus benutzt werden, der auch unter Windows 95 funktioniert:
        '** Anfang **
        'Position = 3
        'Do While Position < Size
        '    If (Buffer(Position) And &H80&) = &H0& Then
        '        mText = mText & ChrW$(Buffer(Position) And &H7F&)
        '        Position = Position + 1
        '    ElseIf (Buffer(Position) And &HE0&) = &HC0& Then
        '        If (Buffer(Position + 1) And &HC0&) = &H80& Then
        '            mText = mText & ChrW$(((Buffer(Position) And &H1F&) * &H40&) Or (Buffer(Position + 1) And &H3F&))
        '        Else
        '            MsgBox "Ungülige UTF-8-Zeichensequenz", vbExclamation
        '        End If
        '        Position = Position + 2
        '    ElseIf (Buffer(Position) And &HF0&) = &HE0& Then
        '        If (Buffer(Position + 1) And &HC0&) = &H80& Then
        '            If (Buffer(Position + 2) And &HC0&) = &H80& Then
        '                mText = mText & ChrW$(((Buffer(Position) And &HF&) * &H1000&) Or ((Buffer(Position + 1) And &H3F&) * &H40&) + (Buffer(Position + 2) And &H3F&))
        '            Else
        '                MsgBox "Ungülige UTF-8-Zeichensequenz", vbExclamation
        '            End If
        '        Else
        '            MsgBox "Ungülige UTF-8-Zeichensequenz", vbExclamation
        '        End If
        '        Position = Position + 3
        '    Else
        '        MsgBox "Ungülige UTF-8-Zeichensequenz", vbExclamation
        '        Position = Position + 1
        '    End If
        'Loop
        '** Ende **
    ElseIf (Signature And ESignatureUTF16Mask) = ESignatureUTF16BE Then
        mEncoding = EEncodingUTF16BE
        mText = String$((Size - 2) \ 2, 0)
        For Position = 2 To Size - 1 Step 2
            On Error Resume Next
            MCh(mText, Position \ 2) = (buffer(Position) * &H100&) Or buffer(Position + 1)
            On Error GoTo 0
        Next Position
    ElseIf (Signature And ESignatureUTF16Mask) = ESignatureUTF16LE Then
        mEncoding = EEncodingUTF16LE
        mText = String$((Size - 2) \ 2, 0)
        ByteToPointer StrPtr(mText), buffer(2), Size - 2
    Else
        mEncoding = EEncodingANSI
        mText = String$(Size, 0)
        DecodeString CP_ACP, 0, buffer(0), Size, StrPtr(mText), Size
    End If
End Sub

Public Sub Save()

Dim Handle As Long
Dim Size As Long
Dim buffer() As Byte
Dim Position As Long
'Die Variable „Character“ wird nur für den alternativen UTF-8-Algorithmus gebraucht
Dim Character As Long

Select Case mEncoding
    Case EEncodingANSI
        Size = Len(mText)
        If Size > 0 Then
            ReDim buffer(0 To Size - 1)
            EncodeStringByte CP_ACP, 0, StrPtr(mText), Size, buffer(0), Size, 0, 0
        End If
    Case EEncodingUTF8
        Size = EncodeStringPointer(CP_UTF8, 0, StrPtr(mText), Len(mText), 0, 0, 0, 0) + 3
        ReDim buffer(0 To Size - 1)
        LongToByte buffer(0), ESignatureUTF8, 3
        EncodeStringByte CP_UTF8, 0, StrPtr(mText), Len(mText), buffer(3), Size - 3, 0, 0
        'Unter Windows 95 funktioniert "WideCharToMultiByte" nicht mit UTF-8-Eingaben.
        'Stattdessen kann die gleichnamige Funktion aus der Bibliothek "unicows.dll" (Microsoft Layer for Unicode) benutzt werden.
        'Die Deklaration und der Aufruf der beiden Funktionen ist bis auf die Bibliotheksangabe völlig gleich.
        'Will man die Bibliothek "unicows.dll" nicht mit ausliefern, kann folgender Algorithmus benutzt werden, der auch unter Windows 95 funktioniert:
        '** Anfang **
        'Size = 3
        'ReDim Buffer(0 To 2)
        'LongToByte Buffer(0), ESignatureUTF8, 3
        'For Position = 1 To Len(mText)
        '    Character = MCh(mText, Position)
        '    Select Case Character
        '        Case 0 To &H7F&
        '            ReDim Preserve Buffer(0 To Size)
        '            Buffer(Size) = Character
        '            Size = Size + 1
        '        Case &H80& To &H7FF&
        '            ReDim Preserve Buffer(0 To Size + 1)
        '            Buffer(Size) = &HC0& Or ((Character And &H3FC0&) \ &H40&)
        '            Buffer(Size + 1) = &H80& Or (Character And &H3F&)
        '            Size = Size + 2
        '        Case &H800& To &HFFFF&
        '            ReDim Preserve Buffer(0 To Size + 2)
        '            Buffer(Size) = &HE0& Or ((Character And &HF000&) \ &H1000&)
        '            Buffer(Size + 1) = &H80& Or ((Character And &HFC0&) \ &H40&)
        '            Buffer(Size + 2) = &H80& Or (Character And &H3F&)
        '            Size = Size + 3
        '    End Select
        'Next Position
        '** Ende **
    Case EEncodingUTF16BE
        Size = LenB(mText) + 2
        ReDim buffer(0 To Size - 1)
        LongToByte buffer(0), ESignatureUTF16BE, 2
        For Position = 0 To Size - 3 Step 2
            PointerToByte buffer(Position + 2), StrPtr(mText) + Position + 1, 1
            PointerToByte buffer(Position + 3), StrPtr(mText) + Position, 1
        Next Position
    Case EEncodingUTF16LE
        Size = LenB(mText) + 2
        ReDim buffer(0 To Size - 1)
        LongToByte buffer(0), ESignatureUTF16LE, 2
        PointerToByte buffer(2), StrPtr(mText), Size - 2
    Case EEncodingUTF32BE
        Size = (LenB(mText) * 2) + 4
        ReDim buffer(0 To Size - 1)
        LongToByte buffer(0), ESignatureUTF32BE, 4
        For Position = 0 To (Size \ 2) - 3 Step 2
            PointerToByte buffer((Position * 2) + 6), StrPtr(mText) + Position + 1, 1
            PointerToByte buffer((Position * 2) + 7), StrPtr(mText) + Position, 1
        Next Position
    Case EEncodingUTF32LE
        Size = (LenB(mText) * 2) + 4
        ReDim buffer(0 To Size - 1)
        LongToByte buffer(0), ESignatureUTF32LE, 4
        For Position = 0 To (Size \ 2) - 3 Step 2
            PointerToByte buffer((Position * 2) + 4), StrPtr(mText) + Position, 2
        Next Position
End Select

If Len(mPath) Then
    Handle = FreeFile
    Open mPath For Binary Access Write As #Handle
    If Size > 0 Then Put #Handle, , buffer
    Close #Handle
End If

End Sub

Public Property Get path() As String
    path = mPath
End Property

Public Property Let path(ByRef NewValue As String)
    mPath = NewValue
End Property

Public Property Get Text() As String
    Text = mText
End Property

Public Property Let Text(ByRef NewValue As String)
    mText = NewValue
End Property

Public Property Get Encoding() As EEncodingScheme
    Encoding = mEncoding
End Property

Public Property Let Encoding(ByVal NewValue As EEncodingScheme)
    mEncoding = NewValue
End Property

Public Function GetShortName(ByVal strLongName As String) As String
  Dim ShortName As String
  Dim sLen As Long
 
  ShortName = Space$(256)
  sLen = GetShortPathName(StrPtr(strLongName), _
    StrPtr(ShortName), Len(ShortName))
  ShortName = Left$(ShortName, sLen)
 
  GetShortName = ShortName
End Function


