VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EasyPopupMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' EasyPopupMenu
' Shahin Gasanov
' www.gasanov.net

Option Explicit

Public Enum MF
    MF_STRING = &H0
    MF_GRAYED = &H1
    MF_CHECKED = &H8
    MF_SEPARATOR = &H800
End Enum

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Const MF_POPUP = &H10
Private Const MF_BYPOSITION = &H400

Private Const TPM_RETURNCMD = &H100

Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function CheckMenuRadioItem Lib "user32" (ByVal hMenu As Long, ByVal un1 As Long, ByVal un2 As Long, ByVal un3 As Long, ByVal un4 As Long) As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hWnd As Long, ByVal lprc As Any) As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private hMenu As Long, hSubmenu As Long

Private Sub Class_Initialize()
    CreateMenu
End Sub

Private Sub Class_Terminate()
    DeleteMenu
End Sub

Public Sub CreateMenu()
    If hMenu = 0 Then
        hMenu = CreatePopupMenu
    End If
End Sub

Public Sub AddMenuItem(ByVal ItemCaption As String, ByVal ItemType As MF, ByVal ItemNumber As Long)
    AppendMenu hMenu, ItemType, ItemNumber, ItemCaption
End Sub

Public Sub CreateSubmenu(ByVal ItemCaption As String)
    hSubmenu = CreatePopupMenu
    AppendMenu hMenu, MF_POPUP, hSubmenu, ItemCaption
End Sub

Public Sub AddSubmenuItem(ByVal ItemCaption As String, ByVal ItemType As MF, ByVal ItemNumber As Long)
    AppendMenu hSubmenu, ItemType, ItemNumber, ItemCaption
End Sub

Public Sub CheckRadioItem(ByVal FirstItem As Long, ByVal LastItem As Long, ByVal RadioItem As Long)
    CheckMenuRadioItem hSubmenu, FirstItem, LastItem, RadioItem, MF_BYPOSITION
End Sub

Public Function TrackMenu(ByVal FormHandler As Long) As Long
    Dim pa As POINTAPI
    
    GetCursorPos pa
    TrackMenu = TrackPopupMenu(hMenu, TPM_RETURNCMD, _
    pa.x, pa.y, 0, FormHandler, 0&)
End Function

Public Sub DeleteMenu()
    If hMenu Then
        If DestroyMenu(hMenu) Then
            hMenu = 0
        End If
    End If
End Sub
