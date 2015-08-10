VERSION 5.00
Begin VB.Form frmogl 
   Caption         =   "OGL"
   ClientHeight    =   10785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13725
   LinkTopic       =   "Form1"
   Picture         =   "frmogl.frx":0000
   ScaleHeight     =   10785
   ScaleWidth      =   13725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Back to Menu"
      Height          =   615
      Left            =   5640
      Picture         =   "frmogl.frx":7C6D5
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   10080
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   9855
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmogl.frx":237487
      Top             =   120
      Width           =   13455
   End
End
Attribute VB_Name = "frmogl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdback_Click()
frmmenu.Show 'goes back to the menu
Unload frmogl
End Sub
