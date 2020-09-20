VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GdipTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const API_NULL As Long = 0
Private Const API_FALSE As Long = 0
Private Const API_TRUE As Long = Not API_FALSE

Private Const S_OK As Long = 0

Private Type UUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Enum EncoderParameterValueType
'    EncoderParameterValueTypeByte = 1
'    EncoderParameterValueTypeASCII = 2
'    EncoderParameterValueTypeShort = 3
    EncoderParameterValueTypeLong = 4
'    EncoderParameterValueTypeRational = 5
'    EncoderParameterValueTypeLongRange = 6
'    EncoderParameterValueTypeUndefined = 7
'    EncoderParameterValueTypeRationalRange = 8
End Enum

Private Type EncoderParameter
    GUID As UUID
    NumberOfValues As Long
    Type As EncoderParameterValueType
    pValue As Long
End Type

Private Type EncoderParameters
    Count As Long 'Must always be set to 1 here, we have just one declared below.
    Parameter As EncoderParameter
End Type

Private Declare Function CLSIDFromString Lib "ole32" ( _
    ByVal lpsz As Long, _
    ByRef clsid As UUID) As Long

Private Declare Function GdipCreateBitmapFromStream Lib "gdiplus" ( _
    ByVal Stream As IUnknown, _
    ByRef GdipBitmap As Long) As Long

Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long

Private Declare Function GdipGetImageHeight Lib "gdiplus" ( _
    ByVal GdipImage As Long, _
    ByRef height As Long) As Long

Private Declare Function GdipGetImageThumbnail Lib "gdiplus" ( _
    ByVal GdipImage As Long, _
    ByVal thumbWidth As Long, _
    ByVal thumbHeight As Long, _
    ByRef GdipThumbImage As Long, _
    ByVal fpCallback As Long, _
    ByVal pCallbackData As Long) As Long

Private Declare Function GdipGetImageWidth Lib "gdiplus" ( _
    ByVal GdipImage As Long, _
    ByRef width As Long) As Long

Private Declare Function GdipSaveImageToFile Lib "gdiplus" ( _
    ByVal GdipImage As Long, _
    ByVal pFileName As Long, _
    ByRef clsidEncoder As UUID, _
    ByVal pEncoderParams As Long) As Long

Private Declare Function GdiplusStartup Lib "gdiplus" ( _
    ByRef Token As Long, _
    ByRef InputBuf As GdiplusStartupInput, _
    ByVal pOutputBuf As Long) As Long

Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal Token As Long) As Long

Private GdipToken As Long
Private GdipStarted As Boolean
Private GifEncoderGUID As UUID
Private JpegEncoderGUID As UUID
Private JpegQualityGUID As UUID
Private PngEncoderGUID As UUID

Public Enum PicFileFormats
    PFF_GIF
    PFF_JPEG
    PFF_PNG
End Enum
#If False Then 'Preserve case of identifiers:
Dim PFF_GIF, PFF_JPEG, PFF_PNG
#End If

Public Function PropStream2PicFileScaled( _
    ByVal PropStream As IUnknown, _
    ByVal FileName As String, _
    ByVal PicFileFormat As PicFileFormats, _
    Optional ByVal JpegQuality As Integer = 90, _
    Optional ByVal MaxWidth As Long = 200, _
    Optional ByVal MaxHeight As Long = 160) As Long
    'Returns GdipThumbImage
    
    Dim FmtGUID As UUID
    Dim Params As EncoderParameters
    Dim pParams As Long
    Dim GdipBitmap As Long
    Dim OrigWidth As Long
    Dim OrigHeight As Long
    Dim Factor As Double
    Dim GdipThumbImage As Long
    Dim rc As Long

    If GdipStarted Then
        Select Case PicFileFormat
            Case PFF_GIF
                FmtGUID = GifEncoderGUID
            Case PFF_JPEG
                FmtGUID = JpegEncoderGUID
                With Params
                    .Count = 1
                    With .Parameter
                        .GUID = JpegQualityGUID
                        .NumberOfValues = 1
                        .Type = EncoderParameterValueTypeLong
                        .pValue = VarPtr(JpegQuality)
                    End With
                End With
                pParams = VarPtr(Params)
            Case PFF_PNG
                FmtGUID = PngEncoderGUID
            Case Else
                Exit Function
        End Select
        If GdipCreateBitmapFromStream(PropStream, GdipBitmap) = S_OK Then
            'Maintain aspect ratio:
            If GdipGetImageWidth(GdipBitmap, OrigWidth) = S_OK Then
                If GdipGetImageHeight(GdipBitmap, OrigHeight) = S_OK Then
                    Factor = CDbl(MaxWidth) / CDbl(OrigWidth)
                    If Int(Factor * CDbl(OrigHeight) + 0.5) > MaxHeight Then
                        Factor = CDbl(MaxHeight) / CDbl(OrigHeight)
                    End If
                    rc = GdipGetImageThumbnail(GdipBitmap, _
                                             Int(Factor * CDbl(OrigWidth) + 0.5), _
                                             Int(Factor * CDbl(OrigHeight) + 0.5), _
                                             GdipThumbImage, _
                                             0, _
                                             0)
                    If rc = S_OK Then
                        PropStream2PicFileScaled = _
                            GdipSaveImageToFile(GdipThumbImage, StrPtr(FileName), FmtGUID, pParams) = S_OK
                        GdipDisposeImage GdipThumbImage
                    End If
                End If
            End If
            GdipDisposeImage GdipBitmap 'These are subclassed GDIP Image objects.
            PropStream2PicFileScaled = GdipThumbImage
        End If
    End If
End Function

Private Sub Class_Initialize()
    Dim GdipSI As GdiplusStartupInput
    
    CLSIDFromString StrPtr("{557cf402-1a04-11d3-9a73-0000f81ef32e}"), GifEncoderGUID
    CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), JpegEncoderGUID
    CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), JpegQualityGUID
    CLSIDFromString StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), PngEncoderGUID
    
    GdipSI.GdiplusVersion = 1
    GdipStarted = GdiplusStartup(GdipToken, GdipSI, API_NULL) = S_OK
End Sub

Private Sub Class_Terminate()
    If GdipStarted Then GdiplusShutdown GdipToken
End Sub
