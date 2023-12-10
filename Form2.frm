VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "close"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Tag             =   "Closing the form"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdPRUEBA 
      Caption         =   "PRUEBA"
      Height          =   1455
      Left            =   960
      TabIndex        =   0
      Tag             =   "ES LO QUE HAY,LTBR"
      Top             =   720
      Width           =   2295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    MsgBox Split(cmdClose.Tag, ",")(0)
    Unload Me
End Sub

Private Sub cmdPRUEBA_Click()

    MsgBox Split(cmdPRUEBA.Tag, ",")(0)
    
End Sub


Private Sub Form_Resize()
'In this form, the tag property is the property 1 (0 based) property, so the first property could by the same as before use the clsAnchor
    Static oAnchor As clsAnchor
    If Not oAnchor Is Nothing Then
    Else
        Set oAnchor = New clsAnchor
        'If no more properties are codified in Tag property the next sentence is not needed
        oAnchor.PropertyNumber = 1
        oAnchor.Separator = ","
        oAnchor.Form = Me
    End If
    
    oAnchor.Anchor
    
End Sub


