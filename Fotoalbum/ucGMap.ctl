VERSION 5.00
Begin VB.UserControl ucGMap 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "ucGMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Implementation of the Google-Static-API (limited to 1000 requests per User and Day)
'There's no dependencies to a Browser-Control - just the plain AsynRead-functionality of a VB6-Usercontrol (the only two API-calls are used, to blit with HalfTone-Quality)
'Author: Olaf Schmidt (2012)
'2013... adjustments to the location-search-api, which now requires a new URL:
'        "http://maps.googleapis.com/maps/api/geocode/xml?&sensor=false&address=" & UTF8-encoded-Address

Option Explicit

Public Enum MapType
  mt_roadmap
  mt_satellite
  mt_hybrid
  mt_terrain
End Enum

Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdcDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nDestWidth As Long, ByVal nDestHeight As Long, ByVal hdcSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal hSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

'Event MouseMove(ByVal GMouseCoordLatLng As String)
Event MouseDown(ByVal GMouseCoordLatLng As String, Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(ByVal GMouseCoordLatLng As String, Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(ByVal GMouseCoordLatLng As String, Button As Integer, Shift As Integer, x As Single, y As Single)
Event DblClick(ByVal GMouseCoordLatLng As String)
 
Private Const mSize& = 640 'this is the max (free usable) SquareSize of the GMap-Static-API
Private Const PI# = 3.14159265358979, TwoPI# = 6.28318530717959
Private Const D2RFac# = 1.74532925199433E-02

Private BackBuf As VB.PictureBox
Private mGZoom As Long, mGPoint As String, mMapType As MapType
Private mLat As Single, mLng As Single
Private mPxlX As Long, mPxlY As Long
Private MDownPoint, LastGMouseMovePoint As String, LastGSearchPoint As String

Public Markers As New Collection

Private Sub UserControl_Initialize()
  ScaleMode = vbPixels
  Set BackBuf = Controls.Add("VB.PictureBox", "BackBuf")
  BackBuf.BorderStyle = 0
  BackBuf.AutoRedraw = True
  BackBuf.Move 0, 0, mSize, mSize
  mGPoint = "0,0"
End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
  If AsyncProp.StatusCode <> vbAsyncStatusCodeEndDownloadData Then Exit Sub
  
  Select Case TypeName(AsyncProp.Value)
    Case "Byte()"
      Dim XML As String
       XML = ReadTagContent(StrConv(AsyncProp.Value, vbUnicode), "location")
  
       LastGSearchPoint = ReadTagContent(XML, "lat") & "," & ReadTagContent(XML, "lng")
    Case "Picture"
      If AsyncProp.bytesRead < 8000 Then Exit Sub
      Set BackBuf.Picture = AsyncProp.Value
      UserControl_Paint
  End Select
End Sub

Private Function ReadTagContent(sXML As String, Tag As String) As String
Dim Result As String
  Result = Mid$(sXML, InStr(sXML, "<" & Tag & ">") + Len(Tag) + 2)
  Result = Left$(Result, InStr(Result, "</" & Tag & ">") - 1)
  Result = Replace(Replace(Result, vbCr, ""), vbLf, "")
  ReadTagContent = Trim$(Result)
End Function

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  x = (x / ScaleWidth - 0.5) * mSize
  y = (y / ScaleHeight - 0.5) * mSize
  If Button = 1 Then MDownPoint = Array(CLng(x), CLng(y))
  RaiseEvent MouseDown(LastGMouseMovePoint, Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  x = (x / ScaleWidth - 0.5) * mSize
  y = (y / ScaleHeight - 0.5) * mSize
  LastGMouseMovePoint = Trim(str(PxlYToLat(mPxlY + y))) & "," & Trim(str(PxlXToLng(mPxlX + x)))
    If frmGEOFinden.blnBeginneRechteck = True Then
        RaiseEvent MouseMove(LastGMouseMovePoint, Button, Shift, x, y)                  'Gerbing 05.09.2016
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  x = (x / ScaleWidth - 0.5) * mSize
  y = (y / ScaleHeight - 0.5) * mSize
  If Not IsEmpty(MDownPoint) Then
    Dim dx As Long, dy As Long
    dx = MDownPoint(0) - x: dy = MDownPoint(1) - y
    MDownPoint = Empty
    GPoint = Trim(str(PxlYToLat(mPxlY + dy))) & "," & Trim(str(PxlXToLng(mPxlX + dx)))
  End If
  RaiseEvent MouseUp(LastGMouseMovePoint, Button, Shift, x, y)
End Sub

Private Sub UserControl_DblClick()
  RaiseEvent DblClick(LastGMouseMovePoint)
End Sub

Private Sub UserControl_Resize()
  UserControl_Paint
End Sub

Private Sub UserControl_Paint()
  SetStretchBltMode UserControl.hdc, 4
  StretchBlt hdc, 0, 0, ScaleWidth, ScaleHeight, BackBuf.hdc, 0, 0, BackBuf.Width, BackBuf.Height, vbSrcCopy
End Sub

'Zoom-related Props
Public Property Let GZoom(ByVal NewValue As Long)
  If NewValue < 0 Then NewValue = 0
  If NewValue > 20 Then NewValue = 20
  If mGZoom = NewValue Then Exit Property
  mGZoom = NewValue
  mPxlX = LngToPxlX
  mPxlY = LatToPxlY
  Refresh
End Property

Public Property Get GZoom() As Long
  GZoom = mGZoom
End Property

Public Property Get RealZoom() As Long
  RealZoom = 2 ^ mGZoom
End Property

'Lat,Long "Csv-String-Point"-related Props
Public Property Let GPoint(ByVal NewValue As String)
Dim Sarr() As String
  If mGPoint = NewValue Then Exit Property
  mGPoint = NewValue
  Sarr = Split(NewValue, ",")
  If UBound(Sarr) <> 1 Then Err.Raise vbObjectError, , _
                            "not a valid Lat,Long-Point-Definition"
  mLat = Val(Sarr(0))
  mLng = Val(Sarr(1))

  mPxlX = LngToPxlX
  mPxlY = LatToPxlY
  Refresh
End Property

Public Property Get GPoint() As String
  GPoint = mGPoint
End Property

Public Property Get lat() As Double
  lat = mLat
End Property

Public Property Get Lng() As Double
  Lng = mLng
End Property

'maptype-Props
Public Property Let MapType(ByVal NewValue As MapType)
  If NewValue < 0 Then NewValue = 0
  If NewValue > 3 Then NewValue = 3
  mMapType = NewValue
  Refresh
End Property

Public Property Get MapType() As MapType
  MapType = mMapType
End Property

Public Function GetMapType(Optional MapType) As String
  If IsMissing(MapType) Then MapType = mMapType
  Select Case MapType
    Case mt_roadmap:   GetMapType = "roadmap"
    Case mt_satellite: GetMapType = "satellite"
    Case mt_hybrid:    GetMapType = "hybrid"
    Case mt_terrain:   GetMapType = "terrain"
  End Select
End Function

'all the Pxl to GeoCoord-formulas found in PHP-code from Fabrice Bernhard
Public Function LngToPxlX(Optional Lng, Optional GZoom, Optional GImgWidth) As Long
  If IsMissing(Lng) Then Lng = mLng
  If IsMissing(GZoom) Then GZoom = mGZoom
  If IsMissing(GImgWidth) Then GImgWidth = mSize
  If Lng > 180 Then Lng = 180 Else If Lng < -180 Then Lng = -180
  
  LngToPxlX = (D2RFac * Lng + PI) * 256 / TwoPI * 2 ^ GZoom
End Function

Public Function LatToPxlY(Optional lat, Optional GZoom, Optional GImgHeight) As Long
  If IsMissing(lat) Then lat = mLat
  If IsMissing(GZoom) Then GZoom = mGZoom
  If IsMissing(GImgHeight) Then GImgHeight = mSize
  If lat > 85 Then lat = 85 Else If lat < -85 Then lat = -85
  
  LatToPxlY = (PI - Log(Tan(PI / 4 + D2RFac * lat / 2))) * 256 / TwoPI * 2 ^ GZoom
End Function

Public Function PxlXToLng(Optional PxlX, Optional GZoom, Optional GImgWidth) As Single
  If IsMissing(PxlX) Then PxlX = mPxlX
  If IsMissing(GZoom) Then GZoom = mGZoom
  If IsMissing(GImgWidth) Then GImgWidth = mSize
  PxlX = PxlX / 2 ^ GZoom
  If PxlX > GImgWidth Then PxlX = GImgWidth Else If PxlX < 0 Then PxlX = 0
  
  PxlXToLng = (PxlX / 256 * TwoPI - PI) / D2RFac
End Function
 
Public Function PxlYToLat(Optional PxlY, Optional GZoom, Optional GImgHeight) As Single
  If IsMissing(PxlY) Then PxlY = mPxlY
  If IsMissing(GZoom) Then GZoom = mGZoom
  If IsMissing(GImgHeight) Then GImgHeight = mSize
  PxlY = PxlY / 2 ^ GZoom
  If PxlY > GImgHeight Then PxlY = GImgHeight Else If PxlY < 0 Then PxlY = 0
  
  PxlYToLat = (2 * Atn(Exp(PI - PxlY / 256 * TwoPI)) - PI / 2) / D2RFac
End Function

Public Function FindLatLngPointFromTextLocation(TextLocation As String) As String
Dim ReqURL As String
  ReqURL = "http://maps.googleapis.com/maps/api/geocode/xml?&sensor=false&address=" & UTF8UrlEnc(TextLocation)
  AsyncRead ReqURL, vbAsyncTypeByteArray, CStr(Timer), vbAsyncReadSynchronousDownload
  FindLatLngPointFromTextLocation = LastGSearchPoint
End Function

Public Sub SetCenterToTextLocation(TextLocation As String)
  GPoint = FindLatLngPointFromTextLocation(TextLocation)
End Sub

Public Sub AddMarker(GPosLatLng As String, ByVal Color As Long, MarkerChar As String)
  Markers.Add "&markers=color:0x" & Color2Hex(Color) & "%7Clabel:" & Left$(MarkerChar, 1) & "%7C" & GPosLatLng
End Sub

Public Function Refresh() As Long
Dim ReqURL As String, M
Static Counter As Long
  Counter = Counter + 1
  'ReqURL = "http://maps.googleapis.com/maps/api/staticmap?sensor=false&format=jpg"
  ReqURL = "http://maps.googleapis.com/maps/api/staticmap?sensor=false&format=jpg&key=AIzaSyCDrk4sJVR8GGanURjDYq0L5OGrXhYDrY4"  'Gerbing 29.09.2018
  ReqURL = ReqURL & "&center=" & GPoint
  ReqURL = ReqURL & "&zoom=" & GZoom
  ReqURL = ReqURL & "&size=" & mSize & "x" & mSize
  ReqURL = ReqURL & "&maptype=" & GetMapType
  For Each M In Markers: ReqURL = ReqURL & M: Next
  
  On Error Resume Next
    AsyncRead ReqURL, vbAsyncTypePicture, "C=" & Counter, vbAsyncReadResynchronize
  If Err Then Err.Clear
  Refresh = Counter
End Function
 
'small Helper-Functions
Private Function Color2Hex(Color As Long) As String
  Color2Hex = Right("0" & Hex(Color \ 65536), 2) & _
              Right("0" & Hex(Color \ 256 And 255), 2) & _
              Right("0" & Hex(Color And 255), 2)
End Function

Private Function UTF8UrlEnc(S As String) As String
Dim i As Long, W As Integer
  For i = 1 To Len(S)
    W = AscW(Mid$(S, i, 1))
    If W < 128 Then
      UTF8UrlEnc = UTF8UrlEnc & ChrW$(W)
    ElseIf W < 2048 Then
      UTF8UrlEnc = UTF8UrlEnc & "%" & Hex$(W \ 64 Or 192) & "%" & Hex$(W And 63 Or 128)
    End If
  Next i
End Function



