VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCreatePic 
      Caption         =   "Create Picture Box"
      Height          =   915
      Left            =   0
      TabIndex        =   0
      Top             =   4020
      Width           =   2235
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Pic1


Private Sub cmdCreatePic_Click()
'these stay the same each time you access the code so they are just Dims
Dim H As Integer
Dim W As Integer
'these don't stay the same each time you access the code so they are Static
Static I As Integer
Static StrName As String

'position of pic box
W = I + 375
H = I + 1135
'Making the Pic box
  Set Pic1 = Form1.Controls.Add("VB.PICTUREBOX", "A" & StrName, Form1)
    'Comment out the line above and un-comment the line below for a textbox
    'Set Pic1 = Form1.Controls.Add("VB.TEXTBOX", "A" & StrName, Form1)
    
Pic1.Left = W
Pic1.Top = H
Pic1.Height = 375
Pic1.Width = 1135
Pic1.Visible = True

'making the pic box in a different location each time
I = I + 100

'Have to make the name of the pic box different each time other wise you get an error
'Because there are ways you can minipulate the item added at run time using the API
'Which include things like pic1_click() ..and you can respond to a click to the new control
StrName = StrName + "B"

End Sub

'You can add most things at run time e.g text boxes , labels , command buttons , toolbars etc.
'just change the line of code saying -
' Pic1 = Form1.Controls.Add("VB.PICTUREBOX", "A" & StrName, Form1)
'Change the ("VB.PICTUREBOX" to some like VB.TEXTBOX or VB.COMMANDBUTTON

