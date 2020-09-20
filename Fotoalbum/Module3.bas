Attribute VB_Name = "Module3"
Option Explicit
 
' benötigte API-Deklarationen                                                                   'Gerbing 15.02.2014
Private Declare Function GradientFillRect Lib "msimg32" _
  Alias "GradientFill" ( _
  ByVal hdc As Long, _
  pVertex As TRIVERTEX, _
  ByVal dwNumVertex As Long, _
  pMesh As GRADIENT_RECT, _
  ByVal dwNumMesh As Long, _
  ByVal dwMode As Long) As Long
 
Private Declare Function GetSysColor Lib "user32" ( _
  ByVal nIndex As Long) As Long
 
Private Declare Function GetClientRect Lib "user32" ( _
  ByVal hwnd As Long, _
  lpRect As RECT) As Long
 
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
 
Private Type GRADIENT_RECT
  UpperLeft As Long
  LowerRight As Long
End Type
 
Private Type TRIVERTEX
  x As Long
  y As Long
  Red As Integer
  Green As Integer
  Blue As Integer
  Alpha As Integer
End Type
 

Public Sub MakeGradient(obj As Object, _
  ByVal ColorFrom As Long, _
  ByVal ColorTo As Long, _
  Optional ByVal nDirection As eGradientDirection)
 
  ' Farbverlauf erzeugen
 
  Dim oRect As RECT
  Dim gRect As GRADIENT_RECT
  Dim oVertex(0 To 1) As TRIVERTEX
 
  ' Prüfen auf Systemfarb-Konstanten und ggf. umwandeln
  If (ColorFrom And &HFF000000) = &H80000000 Then ColorFrom = GetSysColor(ColorFrom And &HFFFFFF)
  If (ColorTo And &HFF000000) = &H80000000 Then ColorTo = GetSysColor(ColorTo And &HFFFFFF)
 
  ' Größe des Objekt-Innenbereichs ermittlen
  GetClientRect obj.hwnd, oRect
 
  ' Linke obere Ecke des Rechtecks
  With oVertex(0)
    .x = 0
    .y = 0
    .Red = sShort((ColorFrom And &HFF&) * 256)
    .Green = sShort((ColorFrom \ &H100& And &HFF&) * 256)
    .Blue = sShort((ColorFrom \ &H10000 And &HFF&) * 256)
    .Alpha = 0
  End With
 
  ' rechte untere Ecke des Rechtecks
  With oVertex(1)
    .x = oRect.Right
    .y = oRect.Bottom
    .Red = sShort((ColorTo And &HFF&) * 256)
    .Green = sShort((ColorTo \ &H100& And &HFF&) * 256)
    .Blue = sShort((ColorTo \ &H10000 And &HFF&) * 256)
    .Alpha = 0
  End With
 
  ' Farbverlauf erstellen
  gRect.UpperLeft = 0
  gRect.LowerRight = 1
  Call GradientFillRect(obj.hdc, oVertex(0), 2, gRect, 1, nDirection)
End Sub

Private Function sShort(ByVal nValue As Long) As Integer
  ' Hilfsfunktion: Umwandeln eines Long-Wertes nach SignedShort
  If nValue < 32768 Then
    sShort = CInt(nValue)
  Else
    sShort = CInt(nValue - &H10000)
  End If
End Function


