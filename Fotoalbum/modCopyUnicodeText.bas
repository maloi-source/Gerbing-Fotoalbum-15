Attribute VB_Name = "modCopyUnicodeText"
Option Explicit

Private Declare Sub RtlMoveMemory Lib "kernel32.dll" ( _
                    ByVal Destination As Long, _
                    ByVal Source As Long, _
                    ByVal Length As Long)
                    
Private Declare Function GlobalAlloc Lib "kernel32.dll" ( _
                         ByVal uFlags As Long, _
                         ByVal dwBytes As Long) As Long
                         
Private Declare Function GlobalSize Lib "kernel32.dll" ( _
                         ByVal hMem As Long) As Long
                         
Private Declare Function GlobalLock Lib "kernel32.dll" ( _
                         ByVal hMem As Long) As Long
                         
Private Declare Function GlobalUnlock Lib "kernel32.dll" ( _
                         ByVal hMem As Long) As Long
                         
Private Declare Function OpenClipboard Lib "user32.dll" ( _
                         ByVal hWndNewOwner As Long) As Long
                         
Private Declare Function EmptyClipboard Lib "user32.dll" () As Long

Private Declare Function GetClipboardData Lib "user32.dll" ( _
                         ByVal uFormat As Long) As Long
                         
Private Declare Function SetClipboardData Lib "user32.dll" ( _
                         ByVal uFormat As Long, _
                         ByVal hMem As Long) As Long
                         
Private Declare Function CloseClipboard Lib "user32.dll" () As Long

Private Const GMEM_MOVEABLE As Long = &H2&
Private Const CF_UNICODETEXT As Long = 13

Public Function ClipboardGetText(ByVal Window As Long) As String
    Dim Memory As Long
    Dim Size As Long
    Dim Pointer As Long
    
    If OpenClipboard(Window) Then
        Memory = GetClipboardData(CF_UNICODETEXT)
        
        If Memory Then
            Size = GlobalSize(Memory)
            
            If Size Then
                Pointer = GlobalLock(Memory)
                
                If Pointer Then
                    ClipboardGetText = Space$((Size \ 2) - 1)
                    RtlMoveMemory StrPtr(ClipboardGetText), Pointer, Size
                    GlobalUnlock Memory
                End If
            End If
        End If
        
        CloseClipboard
    End If
End Function

Public Function ClipboardSetText(ByVal Window As Long, ByRef Text As String, _
    Optional ByVal EmptyBefore As Boolean = True) As Boolean
    
    Dim Size As Long
    Dim Memory As Long
    Dim Pointer As Long
    
    If OpenClipboard(Window) Then
        If EmptyBefore Then EmptyClipboard
        
        Size = LenB(Text) + 2
        Memory = GlobalAlloc(GMEM_MOVEABLE, Size)
        
        If Memory Then
            Pointer = GlobalLock(Memory)
            
            If Pointer Then
                RtlMoveMemory Pointer, StrPtr(Text), Size
                GlobalUnlock Memory
                
                If SetClipboardData(CF_UNICODETEXT, Memory) Then _
                    ClipboardSetText = True
            End If
        End If
        
        CloseClipboard
    End If
End Function


