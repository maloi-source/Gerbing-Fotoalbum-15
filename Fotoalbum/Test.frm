VERSION 5.00
Begin VB.Form Test 
   Caption         =   "Test"
   ClientHeight    =   6888
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11136
   LinkTopic       =   "Form2"
   ScaleHeight     =   6888
   ScaleWidth      =   11136
   StartUpPosition =   3  'Windows-Standard
End
Attribute VB_Name = "Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim result As Variant
Private Sub Form_Load()

    'result = Tan(27.92)
    If isAdmin = True Then
        MsgBox "user hat Administratorrechte"
    Else
        MsgBox "user hat keine Administratorrechte"
    End If
End Sub
