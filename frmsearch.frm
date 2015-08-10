VERSION 5.00
Begin VB.Form frmsearch 
   Caption         =   "search"
   ClientHeight    =   10770
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13470
   LinkTopic       =   "Form1"
   Picture         =   "frmsearch.frx":0000
   ScaleHeight     =   10770
   ScaleWidth      =   13470
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ListBox lstlevels 
      Height          =   4155
      ItemData        =   "frmsearch.frx":57F06
      Left            =   11160
      List            =   "frmsearch.frx":57F08
      TabIndex        =   11
      Top             =   840
      Width           =   1095
   End
   Begin VB.ListBox lstnames 
      Height          =   4155
      ItemData        =   "frmsearch.frx":57F0A
      Left            =   9000
      List            =   "frmsearch.frx":57F0C
      TabIndex        =   10
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "Go to Character Editor"
      Height          =   615
      Left            =   4680
      Picture         =   "frmsearch.frx":57F0E
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdnum 
      Caption         =   "Sort Characters by Level"
      Height          =   615
      Left            =   6600
      Picture         =   "frmsearch.frx":212CC0
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmdalph 
      Caption         =   "Sort Characters by Name (1-9, A-Z, a-z)"
      Height          =   615
      Left            =   4680
      Picture         =   "frmsearch.frx":3CDA72
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmdmenu 
      Caption         =   "Back to Menu"
      Height          =   615
      Left            =   4680
      Picture         =   "frmsearch.frx":588824
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox txtdisplay 
      Height          =   2055
      Left            =   4680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   720
      Width           =   3615
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "Find character folders"
      Height          =   615
      Left            =   6600
      Picture         =   "frmsearch.frx":7435D6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   10575
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmsearch.frx":8FE388
      Top             =   120
      Width           =   4335
   End
   Begin VB.CommandButton cmdend 
      Caption         =   "End Program"
      Height          =   615
      Left            =   6600
      Picture         =   "frmsearch.frx":8FE417
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Name                                        Level   "
      Height          =   255
      Left            =   9000
      TabIndex        =   7
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "List of character folders"
      Height          =   255
      Left            =   5640
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "frmsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim char_lvl(50) As Integer
Dim char_name(50) As String
Dim no_folders As Integer
Dim scroll As Boolean

Private Sub cmdalph_Click()

Dim char(50) As String
Dim list(50) As String
Dim outer As Integer
Dim inner As Integer
Dim min As Integer
Dim counter As Integer
Dim initial As String
Dim initi As String

For counter = 1 To no_folders 'preserves the char_name() array, so that the sort can be done multiple times.
 char(counter) = char_name(counter)
Next counter

'clears both lists
lstlevels.Clear
lstnames.Clear

For outer = 1 To no_folders
 min = outer

 For inner = 1 To no_folders

'get first letter, and convert it into ascii. if the letter is the same, get the next.
 Dim asc_initial As Integer
 Dim asc_initi As Integer
 Dim len1 As Integer
 Dim len2 As Integer
 Dim cntr As Integer
    
 len1 = Len(char(inner))
 len2 = Len(char(min))
  If len1 > len2 Then 'finds the shortest name length, and sets it to len1
   len1 = len2
  End If
    
  For cntr = 1 To len1 'gets a character from each word at the same position, and finds its ASCII value.
   initial = Mid(char(inner), cntr, 1)
   initi = Mid(char(min), cntr, 1)
    
   asc_initial = 0
   asc_initial = asc_initial + CStr(Asc(initial))
    
   asc_initi = 0
   asc_initi = asc_initi + CStr(Asc(initi))

   Select Case asc_initial
    Case Is < asc_initi 'asc_initial comes before asc initi, and is therefore nearer the top of the list.
    min = inner
    cntr = len1 + 1
    Case Is > asc_initi
    cntr = len1 + 1
   End Select
    
  Next cntr
 Next inner

 list(outer) = char(min)
 char(min) = "~" 'sets to an invalid digit, so that it will not come up again. this character has a value of 126, one of the last ascii character.

 lstnames.AddItem list(outer) & vbCrLf
 lstlevels.AddItem char_lvl(min) & vbCrLf

Next outer

End Sub

Private Sub cmdedit_Click()
frmeditor.Show 'opens up the character editor
End Sub

Private Sub cmdend_Click()
End 'ends the program
End Sub

Private Sub cmdmenu_Click()
frmmenu.Show 'goes back to the menu
End Sub

Private Sub cmdnum_Click()

Dim char(50) As Integer
Dim list(50) As Integer
Dim outer As Integer
Dim inner As Integer
Dim min As Integer
Dim counter As Integer


For counter = 1 To no_folders 'preserves the char_lvl() array, so that the sort can be done multiple times.
 char(counter) = char_lvl(counter)
Next counter

'clear the columns for names and levels
lstlevels.Clear
lstnames.Clear


For outer = 1 To no_folders 'uses a selection sort with two lists
 min = outer

 For inner = 1 To no_folders

  If char(inner) < char(min) And char(inner) <> 100 Then 'finds the smallest number.if it is equal to the invalid number, then it is not used
   min = inner
  End If

 Next inner

 list(outer) = char(min)
 char(min) = "100" 'sets to an invalid number, so that it will not come up again

 lstlevels.AddItem list(outer) & vbCrLf
 lstnames.AddItem char_name(min) & vbCrLf

Next outer

End Sub

Private Sub cmdsearch_Click()

'should check d:\character_creator, and list all folders
Dim start As String
Dim root As String
Text1.Text = ""
root = InputBox("Enter the letter of the directory that your characters are saved to. i.e, C;D;X.") 'get root, and validate
    
Do
 If Dir(root & ":", vbDirectory) = "" Then
  root = InputBox("Error, please enter a valid root directory.")
  Else
 End If
Loop Until Dir(root & ":", vbDirectory) <> ""
     
start = root & ":\character_creator" 'start folder containing subfolders
Call ListFolder(start)
End Sub
    
Sub ListFolder(startfolder As String)
Dim FSys As New FileSystemObject
Dim folder As folder
Dim subfolder As folder
Dim counter As Integer
Dim char_info(2, 50) As String
    
'clears the displays
Text1.Text = ""
txtdisplay.Text = ""
lstnames.Clear
lstlevels.Clear
    
Set folder = FSys.GetFolder(startfolder)
For Each subfolder In folder.SubFolders
 DoEvents
 counter = counter + 1
 Debug.Print subfolder 'displays the character folder
 txtdisplay = txtdisplay & subfolder & vbCrLf
    
 Call get_info(subfolder, char_info(), counter)
    
Next subfolder
Set folder = Nothing
MsgBox "Number of subfolders in " & startfolder & " :- " & counter & "/50" 'shows number of characters being stored
no_folders = counter
    
End Sub

Sub get_info(ByVal subfolder As folder, ByRef char_info() As String, ByRef counter As Integer)
    
Dim path As String
Dim info As String
Dim x As String
Dim list(10) As String
Dim star_pos(4) As Integer
Dim length As Integer
Dim char As String
    
path = subfolder & "\txtmisc"
    
Open path For Input As #1
 char = Input$(LOF(1), 1)
Close #1
 
info = char
    
Text1.Text = Text1.Text & vbCrLf & char
    
star_pos(1) = InStr(1, info, "*") 'gets star in front of name
star_pos(2) = InStr(star_pos(1) + 1, info, "*") 'finds star at end of name
length = star_pos(2) - star_pos(1) 'finds length of name
 
char_info(1, counter) = Mid(info, star_pos(1) + 1, length - 1) 'gets name
    
star_pos(3) = InStr(star_pos(2) + 1, info, "*") 'finds star in front of level
star_pos(4) = InStr(star_pos(3) + 1, info, "*") 'finds star at end of level
length = star_pos(4) - star_pos(3) 'finds length of level
    
char_info(2, counter) = Mid(info, star_pos(3) + 1, length - 1) 'gets level
    
'adds name and level to the list
lstnames.AddItem char_info(1, counter)
lstlevels.AddItem char_info(2, counter)
    
char_lvl(counter) = char_info(2, counter)
char_name(counter) = char_info(1, counter)
End Sub
    
Private Sub lstnames_Scroll() 'makes sure that both lists are kept at the same level when scrolled.
If Not scroll Then
    scroll = True
    lstlevels.TopIndex = lstnames.TopIndex
    scroll = False
End If
End Sub

Private Sub lstlevels_Scroll() 'makes sure that both lists are kept at the same level when scrolled.
If Not scroll Then
    scroll = True
    lstnames.TopIndex = lstlevels.TopIndex
    scroll = False
End If
End Sub
