VERSION 5.00
Begin VB.Form frmmenu 
   Caption         =   "menu"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13785
   LinkTopic       =   "Form1"
   Picture         =   "frmmenu.frx":0000
   ScaleHeight     =   5025
   ScaleWidth      =   13785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_ogl 
      Caption         =   "Open Game License"
      Height          =   975
      Left            =   3480
      Picture         =   "frmmenu.frx":573FE
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3840
      Width           =   6495
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "Search Characters"
      Height          =   615
      Left            =   6840
      Picture         =   "frmmenu.frx":2121B0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdopenfile 
      Caption         =   "Read/Edit Character"
      Height          =   615
      Left            =   5160
      Picture         =   "frmmenu.frx":3CCF62
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   645
      Left            =   3600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmmenu.frx":587D14
      Top             =   1800
      Width           =   6255
   End
   Begin VB.CommandButton cmdend 
      Caption         =   "End Program"
      Height          =   615
      Left            =   8400
      Picture         =   "frmmenu.frx":587DA0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdcreate 
      Caption         =   "Create Character"
      Height          =   615
      Left            =   3480
      Picture         =   "frmmenu.frx":5D9556
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Naruto D20 information valid as of May 2012."
      Height          =   255
      Left            =   5160
      TabIndex        =   7
      Top             =   2520
      Width           =   3255
   End
   Begin VB.Label Title 
      Caption         =   "Character Creator v1.0. Copyright ©2012, Thomas Robertson."
      Height          =   255
      Left            =   4440
      TabIndex        =   6
      Top             =   1440
      Width           =   4575
   End
End
Attribute VB_Name = "frmmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_ogl_Click()
frmogl.Show 'opens up the OGL page
End Sub

Private Sub cmdcreate_Click()
frmcharsheet.Show 'brings up the character creator
Unload frmmenu 'unloads the menu form, as its unlikely to be used again.
End Sub

Private Sub cmdend_Click()
End 'ends the program
End Sub

Private Sub Cmdopenfile_Click()
frmeditor.Show 'opens up the character editor
Unload frmmenu 'unloads the menu form, as its unlikely to be used again.
End Sub

Private Sub cmdsearch_Click()
frmsearch.Show 'opens up the character search
Unload frmmenu 'unloads the menu form, as its unlikely to be used again.
End Sub
