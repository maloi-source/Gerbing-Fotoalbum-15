VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GdipLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'Uses GDI+ in order to support PNG and TIFF images as well as GIF, JPEG,
'and BMP.
'
'All failures return Nothing.  This includes the result of any calls made
'if GDI+ failed to start during Class_Initialize.
'

Private Const API_NULL As Long = 0
Private Const API_FALSE As Long = 0
Private Const API_TRUE As Long = Not API_FALSE

Private Const GDIP_OK As Long = 0

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Declare Function GdiplusStartup Lib "gdiplus" ( _
    ByRef Token As Long, _
    ByRef InputBuf As GdiplusStartupInput, _
    ByVal pOutputBuf As Long) As Long

Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal Token As Long) As Long

Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long

Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" ( _
    ByVal bitmap As Long, _
    ByRef hbmReturn As Long, _
    ByVal background As Long) As Long

Private Declare Function GdipCreateBitmapFromFile Lib "gdiplus" ( _
    ByVal pFileName As Long, _
    ByRef bitmap As Long) As Long

Private Declare Function GdipCreateBitmapFromStream Lib "gdiplus" ( _
    ByVal Stream As IUnknown, _
    ByRef bitmap As Long) As Long

Private Type PICTDESC
    cbSizeOfStruct As Long
    picType As Long
    hBitmap As Long
    hpal As Long
End Type

Private Declare Function OleCreatePictureIndirect Lib "oleaut32" ( _
    ByRef PictDescArg As PICTDESC, _
    ByRef riid As Any, _
    ByVal fOwn As Long, _
    ByRef lpvObj As Object) As Long

Private GdipToken As Long
Private GdipStarted As Boolean

Public Function LoadPictureFile(ByVal File As String) As StdPicture
    'Returns Nothing on failure.
    Dim GdipBitmap As Long
    Dim hBitmap As Long 'GDI Bitmap handle.
   
    If GdipStarted Then
        If GdipCreateBitmapFromFile(StrPtr(File), GdipBitmap) = GDIP_OK Then
            If GdipCreateHBITMAPFromBitmap(GdipBitmap, hBitmap, API_NULL) = GDIP_OK Then
                Set LoadPictureFile = BitmapToPicture(hBitmap)
                GdipDisposeImage GdipBitmap 'These are subclassed GDIP Image objects.
            End If
        End If
    End If
End Function

Public Function LoadPictureStream(ByVal ImageStream As IUnknown) As StdPicture
    'Returns Nothing on failure.
    Dim GdipBitmap As Long
    Dim hBitmap As Long 'GDI Bitmap handle.

    If GdipStarted Then
        If GdipCreateBitmapFromStream(ImageStream, GdipBitmap) = GDIP_OK Then
            If GdipCreateHBITMAPFromBitmap(GdipBitmap, hBitmap, API_NULL) = GDIP_OK Then
                Set LoadPictureStream = BitmapToPicture(hBitmap)
                GdipDisposeImage GdipBitmap 'These are subclassed GDIP Image objects.
            End If
        End If
    End If
End Function

Private Function BitmapToPicture(ByVal hBitmap As Long) As StdPicture
    Dim PD As PICTDESC
    Dim IID_IPicture(3) As Long

    With PD
       .cbSizeOfStruct = Len(PD)
       .picType = vbPicTypeBitmap
       .hBitmap = hBitmap
       '.hPal = 0
    End With
    IID_IPicture(0) = &H7BF80980
    IID_IPicture(1) = &H101ABF32
    IID_IPicture(2) = &HAA00BB8B
    IID_IPicture(3) = &HAB0C3000
    OleCreatePictureIndirect PD, IID_IPicture(0), API_TRUE, BitmapToPicture
End Function

Private Sub Class_Initialize()
    Dim GdipSI As GdiplusStartupInput
    
    GdipSI.GdiplusVersion = 1
    GdipStarted = GdiplusStartup(GdipToken, GdipSI, API_NULL) = GDIP_OK
End Sub

Private Sub Class_Terminate()
    If GdipStarted Then GdiplusShutdown GdipToken
End Sub
