VERSION 5.00
Begin VB.Form frmeditor 
   Caption         =   "character editor"
   ClientHeight    =   12750
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16320
   LinkTopic       =   "Form1"
   Picture         =   "frmeditor.frx":0000
   ScaleHeight     =   12750
   ScaleWidth      =   16320
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtspells 
      Height          =   2535
      Left            =   5520
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   3120
      Width           =   3135
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "Delete Character"
      Height          =   615
      Left            =   4440
      Picture         =   "frmeditor.frx":8DB1D
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox txtmods 
      Height          =   2295
      Left            =   3240
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtskills 
      Height          =   6255
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   6240
      Width           =   3255
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save Character"
      Height          =   615
      Left            =   3480
      Picture         =   "frmeditor.frx":2488CF
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdload 
      Caption         =   "Load Character"
      Height          =   615
      Left            =   2520
      Picture         =   "frmeditor.frx":403681
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox txtfeats 
      Height          =   6255
      Left            =   3600
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   6240
      Width           =   3615
   End
   Begin VB.TextBox txtspecial 
      Height          =   12135
      Left            =   9000
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   480
      Width           =   7095
   End
   Begin VB.TextBox txtdisplay 
      Height          =   2295
      Left            =   4800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "frmeditor.frx":5BE433
      Top             =   480
      Width           =   3975
   End
   Begin VB.CommandButton cmdmenu 
      Caption         =   "Back to Menu"
      Height          =   615
      Left            =   3000
      Picture         =   "frmeditor.frx":5BE528
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdend 
      Caption         =   "End Program"
      Height          =   615
      Left            =   3960
      Picture         =   "frmeditor.frx":7792DA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox txtstats 
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox txtmisc 
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label6 
      Caption         =   "Special Abilities"
      Height          =   255
      Left            =   11880
      TabIndex        =   20
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Spells/Day"
      Height          =   255
      Left            =   6600
      TabIndex        =   19
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Feats"
      Height          =   255
      Left            =   5160
      TabIndex        =   18
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "Skills"
      Height          =   255
      Left            =   1560
      TabIndex        =   17
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Stats/Abilities"
      Height          =   255
      Left            =   720
      TabIndex        =   16
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Information for Making Character"
      Height          =   255
      Left            =   5520
      TabIndex        =   15
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Combat Stats"
      Height          =   255
      Left            =   3480
      TabIndex        =   14
      Top             =   240
      Width           =   975
   End
   Begin VB.Label label1 
      Caption         =   "Misc Info"
      Height          =   255
      Left            =   1320
      TabIndex        =   13
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmeditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdmenu_Click()
frmmenu.Show 'goes back to menu
End Sub

Private Sub cmdend_Click()
End 'end program
End Sub

Private Sub cmdload_Click()
Dim root As String
Dim name As String
Dim namelen As Integer
Dim chrname(50) As String
Dim count As Integer
Dim x As Integer
Dim y As Integer
Dim displaylen As Integer
Dim start As String

root = InputBox("Enter the location of the folders, e.g. C;D;N.")

Do
If Dir(root & ":", vbDirectory) = "" Then
root = InputBox("Error, please enter a valid root directory.")
Else
End If
Loop Until Dir(root & ":", vbDirectory) <> ""

'should check :\character_creator, and list all folders

txtdisplay.Text = ""
start = root & ":\character_creator" 'start folder containing subfolders
Call ListFolder(start)

name = InputBox("Enter the name of the character that you wish to load.")

For counter = 1 To 7 'loads the 7 text boxes containing charcter information

 Select Case counter
 Case Is = "1"
 Filename = root & ":\character_creator\" & name & "\txtmisc"
    Open Filename For Input As #1
 txtmisc.Text = Input$(LOF(1), 1)
 Close #1
 
 Case Is = "2"
 Filename = root & ":\character_creator\" & name & "\txtstats"
 Open Filename For Input As #2
 txtstats.Text = Input$(LOF(2), 2)
 Close #2
 
Case Is = "3"
Filename = root & ":\character_creator\" & name & "\txtspecial"
Open Filename For Input As #3
txtspecial.Text = Input$(LOF(3), 3)
 Close #3
 
Case Is = "4"
Filename = root & ":\character_creator\" & name & "\txtfeats"
Open Filename For Input As #4
txtfeats.Text = Input$(LOF(4), 4)
 Close #4
 
Case Is = "5"
Filename = root & ":\character_creator\" & name & "\txtskills"
Open Filename For Input As #5
 txtskills.Text = Input$(LOF(5), 5)
 Close #5
 
 Case Is = "6"
Filename = root & ":\character_creator\" & name & "\txtmods"
Open Filename For Input As #6
 txtmods.Text = Input$(LOF(6), 6)
 Close #6
  
  Case Is = "7"
Filename = root & ":\character_creator\" & name & "\txtspells"
Open Filename For Input As #7
 txtspells.Text = Input$(LOF(7), 7)
 Close #7
  End Select
 Next counter
End Sub

Sub ListFolder(startfolder As String) 'finds the location and name of all character folders
Dim FSys As New FileSystemObject
Dim FSfold As folder
Dim subfolder As folder
Dim counter As Integer
Dim char_info(50) As String
    
Set FSfold = FSys.GetFolder(startfolder) 'gets the initial folder
For Each subfolder In FSfold.SubFolders 'goes through every folder inside the initial folder
DoEvents
counter = counter + 1
Debug.Print subfolder
txtdisplay = txtdisplay & subfolder & vbCrLf 'displays the folder location
Next subfolder
Set FSfold = Nothing
End Sub

Private Sub cmddelete_Click()
Dim character As String
Dim root As String
Dim Response As String

root = InputBox$("Enter the letter of the directory the character is saved in, i.e. C;D;N.")

Do
If Dir(root & ":", vbDirectory) = "" Then
root = InputBox("Error, please enter a valid root directory.")
Else
End If
Loop Until Dir(root & ":", vbDirectory) <> ""

character = InputBox("Enter the name of the character that you wish to delete.")

 Response = MsgBox("Are you sure you want to delete this character", vbYesNo)
 If Response = vbYes Then
  Kill (root & ":/character_creator/" & character) 'deletes the character folder, and all information inside it
 End If
End Sub

Private Sub CmdSave_Click()
Dim counter As Integer
Dim root As String
Dim name As String
Dim Filename As String

MsgBox ("The character should be saved where it was loaded from to prevent errors")
name = InputBox$("Enter your characters name.")
root = InputBox$("Enter the letter of the directory to save in, i.e. C;D;N. the character will then be saved as X:\character_creator\character name\txtmisc")

Do
If Dir(root & ":", vbDirectory) = "" Then
root = InputBox("Error, please enter a valid root directory.")
Else
End If
Loop Until Dir(root & ":", vbDirectory) <> ""

For counter = 1 To 7

 Select Case counter
 Case Is = "1"
 Filename = root & ":\character_creator\" & name & "\txtmisc"
  Open Filename For Output As #1
 Print #1, txtmisc.Text
 Close #1
 
 Case Is = "2"
 Filename = root & ":\character_creator\" & name & "\txtstats"
 Open Filename For Output As #2
 Print #2, txtstats.Text
 Close #2
 
Case Is = "3"
Filename = root & ":\character_creator\" & name & "\txtspecial"
Open Filename For Output As #3
Print #3, txtspecial.Text
 Close #3
 
Case Is = "4"
Filename = root & ":\character_creator\" & name & "\txtfeats"
Open Filename For Output As #4
Print #4, txtfeats.Text
 Close #4
 
Case Is = "5"
Filename = root & ":\character_creator\" & name & "\txtskills"
Open Filename For Output As #5
 Print #5, txtskills.Text
 Close #5
 
 Case Is = "6"
Filename = root & ":\character_creator\" & name & "\txtmods"
Open Filename For Output As #6
 Print #6, txtmods.Text
 Close #6
  
  Case Is = "7"
Filename = root & ":\character_creator\" & name & "\txtspells"
Open Filename For Output As #7
 Print #7, txtspells.Text
 Close #7
  End Select
 Next counter
End Sub
