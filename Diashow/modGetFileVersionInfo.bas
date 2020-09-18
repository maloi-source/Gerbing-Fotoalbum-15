Attribute VB_Name = "modGetFileVersionInfo"
Option Explicit

Private Const MAX_PATH As Long = 260

Private Type VS_FIXEDFILEINFO
    dwSignature         As Long
    dwStrucVersion      As Long
    dwFileVersionMSl    As Integer
    dwFileVersionMSh    As Integer
    dwFileVersionLSl    As Integer
    dwFileVersionLSh    As Integer
    dwProductVersionMSl As Integer
    dwProductVersionMSh As Integer
    dwProductVersionLSl As Integer
    dwProductVersionLSh As Integer
    dwFileFlagsMask     As Long
    dwFileFlags         As Long
    dwFileOS            As Long
    dwFileType          As Long
    dwFileSubtype       As Long
    dwFileDateMS        As Long
    dwFileDateLS        As Long
End Type

Private Declare Function GetFileVersionInfoW Lib "version.dll" (ByVal lptstrFilename As Long, ByVal dwHandle As Long, ByVal dwLen As Long, ByRef lpData As Any) As Long
Private Declare Function GetFileVersionInfoSizeW Lib "version.dll" (ByVal lptstrFilename As Long, Optional ByRef lpdwHandle As Long) As Long
Private Declare Function VerQueryValueW Lib "version.dll" (ByRef pBlock As Any, ByVal lpSubBlock As Long, ByRef lplpBuffer As Long, ByRef puLen As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Public Function GetDiashowExeVersion() As String
    Dim sOperaPath As String, sOperaVer As String
    Dim hData As Long, lDataLen As Long, uBuf() As Byte, uVFFI As VS_FIXEDFILEINFO

   'On Error GoTo Error

    sOperaPath = AppPath & "\diashow.exe"

    If LenB(sOperaPath) Then
        lDataLen = GetFileVersionInfoSizeW(StrPtr(sOperaPath))
        If lDataLen Then
            ReDim uBuf(0& To lDataLen - 1&) As Byte

           'Get handle to file props
            If GetFileVersionInfoW(StrPtr(sOperaPath), 0&, lDataLen, uBuf(0&)) Then
                If VerQueryValueW(uBuf(0&), StrPtr("\"), hData, lDataLen) Then
                    CopyMemory uVFFI, ByVal hData, lDataLen

                    With uVFFI
                        sOperaVer = Format$(.dwFileVersionMSh, "00") & "." & _
                                    Format$(.dwFileVersionMSl, "00") & "." & _
                                    Format$(.dwFileVersionLSl, "0000")
                    End With
                End If
            End If
        End If
    End If

    Select Case False
        Case lDataLen > 0&, AscW(sOperaVer) <> 48   'AscW("0") = 48
            GetDiashowExeVersion = "Unable to get diashow.exe version!"
        Case Else
            GetDiashowExeVersion = sOperaVer
    End Select

End Function


