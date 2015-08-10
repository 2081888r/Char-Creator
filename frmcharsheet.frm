VERSION 5.00
Begin VB.Form frmcharsheet 
   Caption         =   "character creator"
   ClientHeight    =   11430
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   16320
   LinkTopic       =   "Form1"
   Picture         =   "frmcharsheet.frx":0000
   ScaleHeight     =   11430
   ScaleWidth      =   16320
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtspells 
      Height          =   2535
      Left            =   5520
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   3120
      Width           =   3135
   End
   Begin VB.TextBox txtmods 
      Height          =   2295
      Left            =   3240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back to Menu"
      Height          =   615
      Left            =   3960
      Picture         =   "frmcharsheet.frx":8DB1D
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox txtskills 
      Height          =   6255
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   6240
      Width           =   3255
   End
   Begin VB.CommandButton cmdeditor 
      Caption         =   "Save Character and go to Editor"
      Height          =   615
      Left            =   3960
      Picture         =   "frmcharsheet.frx":2488CF
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdclass 
      Caption         =   "Choose Classes"
      Height          =   615
      Left            =   2520
      Picture         =   "frmcharsheet.frx":403681
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton cmdmisc 
      Caption         =   "Get Misc Info"
      Height          =   615
      Left            =   2520
      Picture         =   "frmcharsheet.frx":5BE433
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox txtfeats 
      Height          =   6255
      Left            =   3600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   6240
      Width           =   3615
   End
   Begin VB.TextBox txtspecial 
      Height          =   12135
      Left            =   9000
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   480
      Width           =   7095
   End
   Begin VB.TextBox txtdisplay 
      Height          =   2295
      Left            =   4800
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "frmcharsheet.frx":7791E5
      Top             =   480
      Width           =   4095
   End
   Begin VB.CommandButton cmdstart 
      Caption         =   "Create Character"
      Height          =   615
      Left            =   2520
      Picture         =   "frmcharsheet.frx":779238
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdend 
      Caption         =   "End Program"
      Height          =   615
      Left            =   3960
      Picture         =   "frmcharsheet.frx":933FEA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox txtstats 
      Height          =   2055
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox txtmisc 
      Height          =   2295
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label8 
      Caption         =   "Feats"
      Height          =   255
      Left            =   5160
      TabIndex        =   21
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "Skills"
      Height          =   255
      Left            =   1560
      TabIndex        =   20
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "Special Abilities"
      Height          =   255
      Left            =   11880
      TabIndex        =   19
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Information for Making Character"
      Height          =   255
      Left            =   5520
      TabIndex        =   18
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Spells/Day"
      Height          =   255
      Left            =   6600
      TabIndex        =   17
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Stats/Abilities"
      Height          =   255
      Left            =   720
      TabIndex        =   16
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Combat Stats"
      Height          =   255
      Left            =   3480
      TabIndex        =   15
      Top             =   240
      Width           =   975
   End
   Begin VB.Label label1 
      Caption         =   "Misc Info"
      Height          =   255
      Left            =   1320
      TabIndex        =   14
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmcharsheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdback_Click()
frmmenu.Show 'goes back to the menu
End Sub

Private Sub cmdclass_Click()
frmclasses.Show 'opens up the class selector
End Sub

Private Sub cmdend_Click()
End 'ends the program
End Sub

Private Sub cmdmisc_Click()
'this gets all information that does not have a direct impact on the workings of the character

Dim name As String
Dim age As String
Dim height As String
Dim weight As String
Dim skin As String
Dim hair As String
Dim eyes As String

Dim Filename As String
Dim root As String

'a save location is made before a character can be made.
root = InputBox$("Enter the letter of the directory to save in, i.e. C;D;N. The character will then be saved as 'X:\character_creator\' when you choose to save it.")

Do
If Dir(root & ":", vbDirectory) = "" Then 'validates the directory
root = InputBox("Error, please enter a valid root directory.")
Else
End If
Loop Until Dir(root & ":", vbDirectory) <> ""

Filename = root & ":\character_creator\"

If Dir(Filename, vbDirectory) = "" Then 'automatically creates a save folder if one does not already exist.
MkDir Filename
Else
End If

'asks for misc info
name = InputBox("Enter your characters name.")
age = InputBox("Enter your characters age.")
height = InputBox("Enter your characters height. (metres)")
weight = InputBox("Enter your characters weight. (pounds/lbs)")
skin = InputBox("Enter your characters skin colour.")
hair = InputBox("Enter your characters hair colour.")
eyes = InputBox("Enter your characters eye colour.")

'display misc info
txtmisc = "Name" & vbTab & vbTab & "*" & name & "*" & vbCrLf 'put inbetween * for getting name in search form
txtmisc = txtmisc & "Age" & vbTab & vbTab & age & vbCrLf
txtmisc = txtmisc & "Height" & vbTab & vbTab & height & " m" & vbCrLf 'in metres
txtmisc = txtmisc & "Weight" & vbTab & vbTab & weight & " lbs," & vbCrLf 'in pounds
txtmisc = txtmisc & "Skin Colour" & vbTab & skin & vbCrLf
txtmisc = txtmisc & "Hair Colour" & vbTab & hair & vbCrLf
txtmisc = txtmisc & "Eye Colour" & vbTab & eyes & vbCrLf
End Sub

Private Sub cmdstart_Click()

Dim system As String
Dim level As Integer
Dim Flag As Boolean
Dim no_feats As Integer
Dim skill_points As Integer
Dim Filename As String

Call get_system(system) 'get the system that the player will be using

txtspecial = "" 'clears the special ability text

level = InputBox("Enter the level of your character. Level must be between 1 and 20.") 'gets the character level

Do 'verifies that the level of the character is between 1 and 20, and gives them their first feat
Flag = False
If level >= 1 And level <= 20 Then
no_feats = 1
Flag = True
Else
level = InputBox("Please enter a valid level.")
End If
Loop Until Flag = True

Select Case level 'this adds on the base number of feats they get at their level
Case Is >= 18
no_feats = no_feats + 6
Case Is >= 15
no_feats = no_feats + 5
Case Is >= 12
no_feats = no_feats + 4
Case Is >= 9
no_feats = no_feats + 3
Case Is >= 6
no_feats = no_feats + 2
Case Is >= 3
no_feats = no_feats + 1
End Select

Call statblock_generator(level, system, no_feats, skill_points)
'this will make the random rolls, and let the player assign their stats, as well as chosing their race, and starting occupation if appropiate

txtmisc = txtmisc & "Level" & vbTab & vbTab & "*" & level & "*" & vbCrLf 'put inbetween * for search form

MsgBox ("So far you have " & no_feats & " feats.") 'just a note, could be taken out

txtdisplay = ""
End Sub

Sub get_system(ByRef system As String)

txtdisplay = "D&D" & vbCrLf & "D20 Modern" & vbCrLf & "Naruto D20" 'shows the options that the players have of system
system = InputBox("Enter the system that you wish to use. Choices in display are case specific.")

Do 'makes sure that a valid system is chosen
Flag = False
If system = "D&D" Or system = "D20 Modern" Or system = "Naruto D20" Then
Flag = True
Else
system = InputBox("Please enter a valid system. Choices are case specific.")
End If
Loop Until Flag = True
txtdisplay = system

root = InputBox$("Enter the directory that your character is saved in, i.e. D; C; X.")

Do
If Dir(root & ":", vbDirectory) = "" Then 'validates the directory
root = InputBox("Error, please enter a valid root directory.")
Else
End If
Loop Until Dir(root & ":", vbDirectory) <> ""

 Filename = root & ":\character_creator\system" 'saves the variable so that it can be opened up in another form
 Open Filename For Output As #1
 Print #1, system
 Close #1

End Sub

Sub statblock_generator(ByVal level As Integer, ByVal system As String, ByRef no_feats As Integer, ByRef skill_points As Integer)

Dim stat(6) As Integer
Dim strength As Integer
Dim dexterity As Integer
Dim intelligence As Integer
Dim wisdom As Integer
Dim constitution As Integer
Dim charisma As Integer
Dim str_mod As Integer
Dim dex_mod As Integer
Dim int_mod As Integer
Dim wis_mod As Integer
Dim con_mod As Integer
Dim cha_mod As Integer
Dim roll(4) As Integer
Dim counter As Integer
Dim stat_block As String
Dim choice As String

txtdisplay = "1 - 3d6" & vbCrLf & "2 - 4d6, best 3" & vbCrLf & "3 - 4d6, best 3, reroll 1s" & vbCrLf & "4 - 4d6, best 3, reroll 1s & 2s"
choice = InputBox("Select which method of stat-generation you are using from the list. Use the number indicated before the choice.")

Do 'makes sure that a valid generation is chosen
If choice <> 1 And choice <> 2 And choice <> 3 And choice <> 4 Then
choice = InputBox("Select which method of stat-generation you are using from the list. Enter either '1', '2', '3', or '4'.")
End If
Loop Until choice = 1 Or choice = 2 Or choice = 3 Or choice = 4

For counter = 1 To 6 'get 6 stats using random dice rolls
Call get_rolls(choice, roll())
Call calc_stats(choice, roll(), stat(), counter)
Next counter

Call decide_stats(stat(), strength, dexterity, intelligence, wisdom, constitution, charisma, level, system, no_feats, skill_points)
Call calculate_mods(str_mod, strength, dex_mod, dexterity, int_mod, intelligence, wis_mod, wisdom, con_mod, constitution, cha_mod, charisma)
Call display_stats(strength, str_mod, dexterity, dex_mod, intelligence, int_mod, wisdom, wis_mod, constitution, con_mod, charisma, cha_mod, stat_block)

txtstats = stat_block
End Sub

Sub get_rolls(ByVal choice As String, ByRef roll() As Integer)

Dim cntr As Integer

Select Case choice 'selects the chosen rolling method, and gets random six-sided dice 'rolls'
Case Is = "1"
For cntr = 1 To 3
roll(cntr) = Int((6 * Rnd) + 1)
Do
If roll(cntr) < 1 Then
roll(cntr) = Int((6 * Rnd) + 1)
End If
Loop Until roll(cntr) >= 1
Next cntr

Case Is = "2"
For cntr = 1 To 4
roll(cntr) = Int((6 * Rnd) + 1)
Do
If roll(cntr) < 1 Then
roll(cntr) = Int((6 * Rnd) + 1)
End If
Loop Until roll(cntr) >= 1
Next cntr

Case Is = "3"
For cntr = 1 To 4
roll(cntr) = Int((6 * Rnd) + 1)
Do
If roll(cntr) <= 1 Then
roll(cntr) = Int((6 * Rnd) + 1)
End If
Loop Until roll(cntr) > 1 'makes sure that the result of 1 is not given
Next cntr

Case Is = "4"
For cntr = 1 To 4
roll(cntr) = Int((6 * Rnd) + 1)
Do
If roll(cntr) <= 2 Then
roll(cntr) = Int((6 * Rnd) + 1)
End If
Loop Until roll(cntr) > 2 'makes sure that the result of 1 or 2 is not given
Next cntr
End Select

End Sub

Sub calc_stats(ByVal choice As String, ByRef roll() As Integer, ByRef stat() As Integer, ByRef counter As Integer)

Dim cntr As Integer
Dim lowest_roll As Integer

If choice = "1" Then 'get each stat
stat(counter) = roll(1) + roll(2) + roll(3)

Else 'when choice is 2,3,or 4
lowest_roll = roll(1)
For cntr = 1 To 3 'gets the lowest roll, which is taken off the total at the end
If lowest_roll > roll(cntr + 1) Then 'gets the lowest roll
lowest_roll = roll(cntr + 1)
End If

Next cntr
stat(counter) = roll(1) + roll(2) + roll(3) + roll(4) - lowest_roll 'adds all rolls, then takes away the lowest
End If

End Sub

Sub decide_stats(ByRef stat() As Integer, ByRef strength As Integer, ByRef dexterity As Integer, ByRef intelligence As Integer, ByRef wisdom As Integer, ByRef constitution As Integer, ByRef charisma As Integer, ByVal level As Integer, ByVal system As String, ByRef no_feats As Integer, ByRef skill_points As Integer)
Dim letter(6) As String
Dim exclude(5) As String
Dim points As Integer
Dim entered As String
Dim counter As Integer
Dim race As String

txtstats = "a" & vbTab & stat(1) & vbCrLf & "b" & vbTab & stat(2) & vbCrLf & "c" & vbTab & stat(3) & vbCrLf & "d" & vbTab & stat(4) & vbCrLf & "e" & vbTab & stat(5) & vbCrLf & "f" & vbTab & stat(6) & vbCrLf

Do
letter(1) = InputBox("Select the letter of the value that you wish to place in strength.")
Loop Until letter(1) = "a" Or letter(1) = "b" Or letter(1) = "c" Or letter(1) = "d" Or letter(1) = "e" Or letter(1) = "f"

Select Case letter(1)
Case Is = "a"
strength = stat(1)
exclude(1) = "a"
Case Is = "b"
strength = stat(2)
exclude(1) = "b"
Case Is = "c"
strength = stat(3)
exclude(1) = "c"
Case Is = "d"
strength = stat(4)
exclude(1) = "d"
Case Is = "e"
strength = stat(5)
exclude(1) = "e"
Case Is = "f"
strength = stat(6)
exclude(1) = "f"
End Select

Do
Do
'loop until letter =a-f, then check if =exclude. loop again if letter=exclude

letter(2) = InputBox("Select the letter of the value that you wish to place in dexterity. You have already used: " & exclude(1))
If letter(2) = exclude(1) Then
MsgBox ("Error, please re-enter.")
End If
Loop Until letter(2) <> exclude(1)
Loop Until letter(2) = "a" Or letter(2) = "b" Or letter(2) = "c" Or letter(2) = "d" Or letter(2) = "e" Or letter(2) = "f"

Select Case letter(2)
Case Is = "a"
dexterity = stat(1)
exclude(2) = "a"
Case Is = "b"
dexterity = stat(2)
exclude(2) = "b"
Case Is = "c"
dexterity = stat(3)
exclude(2) = "c"
Case Is = "d"
dexterity = stat(4)
exclude(2) = "d"
Case Is = "e"
dexterity = stat(5)
exclude(2) = "e"
Case Is = "f"
dexterity = stat(6)
exclude(2) = "f"
End Select

Do
Do
letter(3) = InputBox("Select the letter of the value that you wish to place in intelligence. You have already used: " & exclude(1) & exclude(2))
If letter(3) = exclude(1) Or letter(3) = exclude(2) Then
MsgBox ("Error, please re-enter.")
End If
Loop Until letter(3) <> exclude(1) And letter(3) <> exclude(2)
Loop Until letter(3) = "a" Or letter(3) = "b" Or letter(3) = "c" Or letter(3) = "d" Or letter(3) = "e" Or letter(3) = "f"

Select Case letter(3)
Case Is = "a"
intelligence = stat(1)
exclude(3) = "a"
Case Is = "b"
intelligence = stat(2)
exclude(3) = "b"
Case Is = "c"
intelligence = stat(3)
exclude(3) = "c"
Case Is = "d"
intelligence = stat(4)
exclude(3) = "d"
Case Is = "e"
intelligence = stat(5)
exclude(3) = "e"
Case Is = "f"
intelligence = stat(6)
exclude(3) = "f"
End Select

Do
Do
letter(4) = InputBox("Select the letter of the value that you wish to place in wisdom. You have already used: " & exclude(1) & exclude(2) & exclude(3))
If letter(4) = exclude(1) Or letter(4) = exclude(2) Or letter(4) = exclude(3) Then
MsgBox ("Error, please re-enter.")
End If
Loop Until letter(4) <> exclude(1) And letter(4) <> exclude(2) And letter(4) <> exclude(3)
Loop Until letter(4) = "a" Or letter(4) = "b" Or letter(4) = "c" Or letter(4) = "d" Or letter(4) = "e" Or letter(4) = "f"

Select Case letter(4)
Case Is = "a"
wisdom = stat(1)
exclude(4) = "a"
Case Is = "b"
wisdom = stat(2)
exclude(4) = "b"
Case Is = "c"
wisdom = stat(3)
exclude(4) = "c"
Case Is = "d"
wisdom = stat(4)
exclude(4) = "d"
Case Is = "e"
wisdom = stat(5)
exclude(4) = "e"
Case Is = "f"
wisdom = stat(6)
exclude(4) = "f"
End Select

Do
Do
letter(5) = InputBox("Select the letter of the value that you wish to place in constitution. You have already used: " & exclude(1) & exclude(2) & exclude(3) & exclude(4))
If letter(5) = exclude(1) Or letter(5) = exclude(2) Or letter(5) = exclude(3) Or letter(5) = exclude(4) Then
MsgBox ("Error, please re-enter.")
End If
Loop Until letter(5) <> exclude(1) And letter(5) <> exclude(2) And letter(5) <> exclude(3) And letter(5) <> exclude(4)
Loop Until letter(5) = "a" Or letter(5) = "b" Or letter(5) = "c" Or letter(5) = "d" Or letter(5) = "e" Or letter(5) = "f"

Select Case letter(5)
Case Is = "a"
constitution = stat(1)
exclude(5) = "a"
Case Is = "b"
constitution = stat(2)
exclude(5) = "b"
Case Is = "c"
constitution = stat(3)
exclude(5) = "c"
Case Is = "d"
constitution = stat(4)
exclude(5) = "d"
Case Is = "e"
constitution = stat(5)
exclude(5) = "e"
Case Is = "f"
constitution = stat(6)
exclude(5) = "f"
End Select

Do
Do
letter(6) = InputBox("Select the letter of the value that you wish to place in charisma. You have already used: " & exclude(1) & exclude(2) & exclude(3) & exclude(4) & exclude(5))
If letter(6) = exclude(1) Or letter(6) = exclude(2) Or letter(6) = exclude(3) Or letter(6) = exclude(4) Or letter(6) = exclude(5) Then
MsgBox ("error, please re-enter")
End If
Loop Until letter(6) <> exclude(1) And letter(6) <> exclude(2) And letter(6) <> exclude(3) And letter(6) <> exclude(4) And letter(6) <> exclude(5)
Loop Until letter(6) = "a" Or letter(6) = "b" Or letter(6) = "c" Or letter(6) = "d" Or letter(6) = "e" Or letter(6) = "f"

Select Case letter(6)
Case Is = "a"
charisma = stat(1)
Case Is = "b"
charisma = stat(2)
Case Is = "c"
charisma = stat(3)
Case Is = "d"
charisma = stat(4)
Case Is = "e"
charisma = stat(5)
Case Is = "f"
charisma = stat(6)
End Select

'gets the number of stat points gained. 1 point is given at 4th, 8th, 12th, 16th and 20th.
Select Case level
Case Is = 20
points = 5
Case Is >= 16
points = 4
Case Is >= 12
points = 3
Case Is >= 8
points = 2
Case Is >= 4
points = 1
End Select

MsgBox ("You have " & points & " stat points to use. Enter the the prefix of the stat you wish to place it in.")

txtstats = "Stat" & vbTab & "Base" & vbCrLf
txtstats = txtstats & "STR" & vbTab & strength & vbCrLf
txtstats = txtstats & "DEX" & vbTab & dexterity & vbCrLf
txtstats = txtstats & "INT" & vbTab & intelligence & vbCrLf
txtstats = txtstats & "WIS" & vbTab & wisdom & vbCrLf
txtstats = txtstats & "CON" & vbTab & constitution & vbCrLf
txtstats = txtstats & "CHA" & vbTab & charisma & vbCrLf

For counter = 1 To points
Do
entered = InputBox("Enter the prefix where you wish to place a point. 'str', 'dex', 'int', 'wis', 'con', or 'cha'. You have " & points + 1 - counter & " left.") '+1 used to correct counter
Loop Until entered = "str" Or entered = "dex" Or entered = "int" Or entered = "wis" Or entered = "con" Or entered = "cha"

Select Case entered 'add on the stat points
Case Is = "str"
strength = strength + 1
Case Is = "dex"
dexterity = dexterity + 1
Case Is = "int"
intelligence = intelligence + 1
Case Is = "wis"
wisdom = wisdom + 1
Case Is = "con"
constitution = constitution + 1
Case Is = "cha"
charisma = charisma + 1
End Select
Next counter

Select Case system 'gets the race of the character, depending on the system used
Case Is = "D&D"
Call dnd_races(race, no_feats, skill_points, level, dexterity, constitution, strength, intelligence, charisma, wisdom)

Case Is = "Naruto D20"
Call naruto_races(race, no_feats, skill_points, level, strength, constitution, dexterity, charisma, wisdom)
Call naruto_occ(no_feats)

Case Is = "D20 Modern"
Call modern_races(race, no_feats, skill_points, level)
Call modern_occ(no_feats)
End Select

root = InputBox$("Enter the letter of the directory your character is saved in, i.e. C;D;N. This is in order for your skills and feats to pass through.")

Do
If Dir(root & ":", vbDirectory) = "" Then 'validates the directory
root = InputBox("Error, please enter a valid root directory.")
Else
End If
Loop Until Dir(root & ":", vbDirectory) <> ""

Filename = root & ":\character_creator\skill"
  Open Filename For Output As #1
 Print #1, skill_points
 Close #1
 
 Filename = root & ":\character_creator\feat"
  Open Filename For Output As #2
 Print #2, no_feats
 Close #2

 Filename = root & ":\character_creator\level"
  Open Filename For Output As #3
 Print #3, level
 Close #3

End Sub

Sub dnd_races(ByRef race As String, ByRef no_feats As Integer, ByRef skill_points As Integer, ByRef level As Integer, ByRef dexterity As Integer, ByRef constitution As Integer, ByRef strength As Integer, ByRef intelligence As Integer, ByRef charisma As Integer, ByRef wisdom As Integer)

txtdisplay = ("Human" & vbTab & "Elf" & vbCrLf & "Half-Elf" & vbTab & "Half-Orc" & vbCrLf & "Gnome" & vbTab & "Dwarf" & vbCrLf & "Halfling")

Do
race = InputBox("Select a race from the names shown.")
Loop Until race = "Human" Or race = "Elf" Or race = "Half-Elf" Or race = "Half-Orc" Or race = "Gnome" Or race = "Dwarf" Or race = "Halfling"

txtmisc = txtmisc & "Race" & vbTab & vbTab & race & vbCrLf

Select Case race 'gets the characters race, and any bonuses or special effects
Case Is = "Human"
no_feats = no_feats + 1
skill_points = skill_points + 3 + 1 * level

Case Is = "Elf"
dexterity = dexterity + 2
constitution = constitution - 2
txtspecial = txtspecial & "Lowlight vision" & vbCrLf & "Immune to sleep effects. +2 save against enchantments." & vbCrLf & "+2 to listen, search and spot checks." & vbCrLf

Case Is = "Half-Elf"
txtspecial = txtspecial & "Lowlight vision" & vbCrLf & "Immune to sleep effects. +2 save against enchantments." & vbCrLf & "+1 to listen, search and spot checks." & vbCrLf & "+2 to diplomacy and gather info checks" & vbCrLf

Case Is = "Half-Orc"
strength = strength + 2
intelligence = intelligence - 2
charisma = charisma - 2
txtspecial = txtspecial & "Darkvision 60'" & vbCrLf & "Orc Blood - for all race related effects, a half orc is considered an orc." & vbCrLf

Case Is = "Gnome"
constitution = constitution + 2
strength = strength - 2
txtspecial = txtspecial & "Speed = 20 feet." & vbCrLf & "lowlight vision" & vbCrLf & "+2 save against illusions. +1 difficulty against illusion spells cast. stackable." & vbCrLf & "+4 dodge against giant type monsters." & vbCrLf & "+1 on attacks rolls against orcs and goblinoids." & vbCrLf & vbCrLf & "+2 on listen and craft(alchemy) checks" & vbCrLf & "1/day - speak w/ burrowing mammal, duration 1 min. If charisma>=10 then also has 1/day dancing lights ,ghost sound, prestidigitation. caster level 1st, save DC 10+cha mod+spell level." & vbCrLf & "Small creature: +1 ac, +1 attack rolls and +4 on hide checks. uses smaller weapons, and has 3/4 carrying limits of a medium creature." & vbCrLf

Case Is = "Dwarf"
constitution = constitution + 2
charisma = charisma - 2
txtspecial = txtspecial & "Speed = 20 feet. Can move at this speed when wearing medium/heavy armour or when carrying medium/heavy loads." & vbCrLf & "Darkvision 60'" & vbCrLf & "+2 save against poisons and spells." & vbCrLf & "+2 on craft and appraise checks related to stone or metal." & vbCrLf & "+4 dodge against giant type monsters." & vbCrLf & "+1 on attacks rolls against orcs and goblinoids." & vbCrLf

Case Is = "Halfling"
dexterity = dexterity + 2
strength = strength - 2
txtspecial = txtspecial & "+2 to climb, jump, move silently, and listen checks." & vbCrLf & "+1 to all saving throws" & vbCrLf & "+2 morale bonus against fear. stacks with +1 for being a halfling." & vbCrLf & "+1 on attacks with throwing weapons and slings." & vbCrLf & "Small creature: +1 ac, +1 attack rolls and +4 on hide checks. uses smaller weapons, and has 3/4 carrying limits of a medium creature" & vbCrLf
End Select

End Sub

Sub naruto_races(ByRef race As String, ByRef no_feats As Integer, ByRef skill_points As Integer, ByRef level As Integer, ByRef strength As Integer, ByRef constitution As Integer, ByRef dexterity As Integer, ByRef charisma As Integer, ByRef wisdom As Integer)

Dim deform(2) As String

txtdisplay = ("Human" & vbTab & "Gigantic" & vbCrLf & "Earth" & vbTab & "Fire" & vbCrLf & "Water" & vbTab & "Wind" & vbCrLf & "Lightning" & vbTab & "Trueblood" & vbCrLf & "Smallfolk" & vbTab & "Monstrous")

Do
race = InputBox("Select a race from the names shown.")
Loop Until race = "Human" Or race = "Gigantic" Or race = "Earth" Or race = "Fire" Or race = "Water" Or race = "Wind" Or race = "Lightning" Or race = "Trueblood" Or race = "Smallfolk" Or race = "Monstrous"

txtmisc = txtmisc & "Race" & vbTab & vbTab & race & vbCrLf

Select Case race
Case Is = "Human"
no_feats = no_feats + 1
skill_points = skill_points + 3 + 1 * level

Case Is = "Gigantic"
level = level - 2
strength = strength + 8
constitution = constitution + 4
dexterity = dexterity - 2
txtspecial = txtspecial & "Level Adjustment of 2(This has been taken off your level)" & vbCrLf & "Large creature: -1 to defense, attack rolls and -4 to hide checks. +4 to grapple checks.uses larger weapons, and has double the carrying capacity of a medium character." & vbCrLf

Case Is = "Fire"
charisma = charisma + 2
strength = strength - 2
txtspecial = txtspecial & "+1 save against fire based attacks" & vbCrLf & "Inspire Courage (Sp): Once per day, as a swift action, he may grant himself and his allies within 30ft a +1 morale bonus to attack rolls, saves and skill checks, and a +4 morale bonus to saves against fear effects for 1 minute or the duration of an encounter (whichever is shorter). Can be used twice/day at 10th, and thrice/day at 20th." & vbCrLf

Case Is = "Earth"
constitution = constitution + 2
wisdom = wisdom - 2
txtspecial = txtspecial & "+1 save against earth based attacks" & vbCrLf & "Tremorsense (Ex): can concentrate for 1 swift action to activate a tremorsense 30 ft. ability once/day. The tremorsense lasts for 1 minute or the duration of an encounter (whichever is shorter). Can be used twice/day at 10th, and thrice/day at 20th." & vbCrLf

Case Is = "Water"
strength = strength + 2
charisma = charisma - 2
txtspecial = txtspecial & "+1 bonus to saves against water based attacks." & vbCrLf & "Can hold his breath twice as Long before suffocating or drowning" & vbCrLf

Case Is = "Wind"
dexterity = dexterity + 2
wisdom = wisdom - 2
txtspecial = txtspecial & "+1 save against wind based attacks" & vbCrLf & "Quickness (Su): May be activated once/day as an instant action. When a Reflex save against an attack, technique or effect requiring a save for half damage is made, he takes no damage on a successful save. This ability is used whether or not the save was successful. Must be declared before rolling the save. Can be used twice/day at 10th, and thrice/day at 20th." & vbCrLf

Case Is = "Lightning"
dexterity = dexterity + 2
wisdom = wisdom - 2
txtspecial = txtspecial & "+1 save against lightning based attacks" & vbCrLf & "Grounded (Su): Can choose to take only half damage from any single electricity-based attacks, so long as he is in contact with the ground, once/day as an instant action. Must be declared before damage from the ability is rolled. Resistance is applied after halving the damage and saves (if allowed any) are rolled. Can be used twice/day at 10th, and thrice/day at 20th." & vbCrLf

Case Is = "Trueblood"
strength = strength + 2
constitution = constitution + 2
charisma = charisma + 2
level = level - 1
txtspecial = txtspecial & "Level Adjustment of 1(This has been taken off your level)" & vbCrLf & "Lowlight vision: can see twice as far as a human in situations of poor illumination. Retain the ability to distinguish colour and detail" & vbCrLf & "darkvision 60': can see without the aid of light up to 60 feet. Black and white vision only, but otherwise as normal sight." & vbCrLf

Case Is = "Smallfolk"
dexterity = dexterity + 2
txtspecial = txtspecial & "+1 to defense, +1 on attack rolls, +4 on hide checks. -4 on grapple checks." & vbCrLf

Case Is = "Monstrous"
strength = strength + 4
dexterity = dexterity - 2
constitution = constitution + 2
charisma = charisma - 2
level = level - 2
txtdisplay = "Natural Attack: Bite; Claws; Gore; Tail whip. Gains the appropriate body part." & vbCrLf & "Advanced Immune System: Gains a +8 bonus to Fortitude saves against poisons and diseases." & vbCrLf & "Amphibious: Able to breathe water and air normally." & vbCrLf & "Powerful Legs: Has a base land speed of 40 feet." & vbCrLf & "Scales: Gains a +4 natural armor bonus to defense." & vbCrLf & "Lithe: Instead of +4 str -2 dex, stat changes are -2 str, +4 dex." & vbCrLf & "Nocturnal: Gain darkvision and lowlight vision out to 60'."

Do
deform(1) = InputBox("Enter the name of the 1st deformity. e.g. - Natural Attack: Bite, Natural Attack: Tail whip, Amphibious, etc.")
Loop Until deform(1) = "Natural Attack: Bite" Or deform(1) = "Natural Attack: Claws" Or deform(1) = "Natural Attack: Gore" Or deform(1) = "Natural Attack: Tail whip" Or deform(1) = "Amphibious" Or deform(1) = "Advanced Immune System" Or deform(1) = "Powerful legs" Or deform(1) = "Scales"

Select Case deform(1)
Case Is = "Natural Attack: Bite"
txtspecial = txtspecial & "Gains 1 bite attack. 1d4+1/2 str mod." & vbCrLf
Case Is = "Natural Attack: Claws"
txtspecial = txtspecial & "Gains 2 claw attacks. 1d4 + str mod." & vbCrLf
Case Is = "Natural Attack: Gore"
txtspecial = txtspecial & "Gains 1 gore attack. 1d6+1/2 str mod." & vbCrLf
Case Is = "Natural Attack: Tail whip"
txtspecial = txtspecial & "Gains 1 tail whip attack. 1d6+1/2 str mod." & vbCrLf
Case Is = "Amphibious"
txtspecial = txtspecial & "Amphibious: gains the ability to breathe under water." & vbCrLf
Case Is = "Advanced Immune System"
txtspecial = txtspecial & "Adv. Immune System: Gains a +8 fort save against poison and disease." & vbCrLf
Case Is = "Powerful Legs"
txtspecial = txtspecial & "Powerful Legs: Base speed incresed by 10ft." & vbCrLf
Case Is = "Scales"
txtspecial = txtspecial & "Scales: gain a +4 to defence. This should be added to what is given." & vbCrLf
Case Is = "Lithe"
txtspecial = txtspecial & "Strength and Dexterity changes are reversed."
strength = strength - 6
dexterity = dexterity + 6
Case Is = "Nocturnal"
txtspecial = txtspecial & "gain lowlight vision and darkvision out to 60'."
End Select

Do
Do
deform(2) = InputBox("Enter the name of the 2nd deformity. e.g. - natural weapon, amphibious, etc. use lowercase. Already chosen: " & deform(1))
Loop Until deform(2) = "Natural Attack: Bite" Or deform(2) = "Natural Attack: Claws" Or deform(2) = "Natural Attack: Gore" Or deform(2) = "Natural Attack: Tail whip" Or deform(2) = "Amphibious" Or deform(2) = "Advanced Immune System" Or deform(2) = "Powerful legs" Or deform(2) = "Scales"
Loop Until deform(2) <> deform(1)

Select Case deform(2)
Case Is = "Natural Attack: Bite"
txtspecial = txtspecial & "Gains 1 bite attack. 1d4+1/2 str mod." & vbCrLf
Case Is = "Natural Attack: Claws"
txtspecial = txtspecial & "Gains 2 claw attacks. 1d4 + str mod." & vbCrLf
Case Is = "Natural Attack: Gore"
txtspecial = txtspecial & "Gains 1 gore attack. 1d6+1/2 str mod." & vbCrLf
Case Is = "Natural Attack: Tail whip"
txtspecial = txtspecial & "Gains 1 tail whip attack. 1d6+1/2 str mod." & vbCrLf
Case Is = "Amphibious"
txtspecial = txtspecial & "Amphibious: gains the ability to breathe under water." & vbCrLf
Case Is = "Advanced Immune System"
txtspecial = txtspecial & "Adv. Immune System: Gains a +8 fort save against poison and disease." & vbCrLf
Case Is = "Powerful Legs"
txtspecial = txtspecial & "Powerful Legs: Base speed incresed by 10ft." & vbCrLf
Case Is = "Scales"
txtspecial = txtspecial & "Scales: gain a +4 to defence. This should be added to what is given." & vbCrLf
Case Is = "Lithe"
txtspecial = txtspecial & "Strength and Dexterity changes are reversed."
strength = strength - 6
dexterity = dexterity + 6
Case Is = "Nocturnal"
txtspecial = txtspecial & "gain lowlight vision and darkvision out to 60'."
End Select

txtspecial = txtspecial & "Level Adjustment of 2(This has been taken off your level)" & vbCrLf
End Select

End Sub

Sub modern_races(ByRef race As String, ByRef no_feats As Integer, ByRef skill_points As Integer, ByRef level As Integer)

txtdisplay = ("Human") 'human is usually the PC race in modern.

Do
race = InputBox("Select a race from the names shown.")
Loop Until race = "Human"

txtmisc = txtmisc & "Race" & vbTab & vbTab & race & vbCrLf

Select Case race
Case Is = "Human"
no_feats = no_feats + 1
skill_points = skill_points + 3 + 1 * level
End Select

End Sub

Sub calculate_mods(ByRef str_mod As Integer, ByRef strength As Integer, ByRef dex_mod As Integer, ByRef dexterity As Integer, ByRef int_mod As Integer, ByRef intelligence As Integer, ByRef wis_mod As Integer, ByRef wisdom As Integer, ByRef con_mod As Integer, ByRef constitution As Integer, ByRef cha_mod As Integer, ByRef charisma As Integer)

str_mod = ((strength - 10) / 2) - 0.25
dex_mod = ((dexterity - 10) / 2) - 0.25
int_mod = ((intelligence - 10) / 2) - 0.25
wis_mod = ((wisdom - 10) / 2) - 0.25
con_mod = ((constitution - 10) / 2) - 0.25
cha_mod = ((charisma - 10) / 2) - 0.25


root = InputBox$("Enter the letter of the directory your character is saved in, i.e. C;D;N. This is in order for your stats to pass through.")
 
  Do
If Dir(root & ":", vbDirectory) = "" Then 'validates the directory
root = InputBox("Error, please enter a valid root directory.")
Else
End If
Loop Until Dir(root & ":", vbDirectory) <> ""
 
 Filename = root & ":\character_creator\strength"
  Open Filename For Output As #6
 Print #6, strength
 Close #6
 
 Filename = root & ":\character_creator\charisma"
  Open Filename For Output As #7
 Print #7, charisma
 Close #7
 
 Filename = root & ":\character_creator\intelligence"
  Open Filename For Output As #8
 Print #8, intelligence
 Close #8
 
 Filename = root & ":\character_creator\constitution"
  Open Filename For Output As #9
 Print #9, constitution
 Close #9
 
 Filename = root & ":\character_creator\dexterity"
  Open Filename For Output As #10
 Print #10, dexterity
 Close #10
 
 Filename = root & ":\character_creator\wisdom"
  Open Filename For Output As #11
 Print #11, wisdom
 Close #11
 
 Filename = root & ":\character_creator\dex"
  Open Filename For Output As #12
 Print #12, dex_mod
 Close #12

 Filename = root & ":\character_creator\int"
  Open Filename For Output As #13
 Print #13, int_mod
 Close #13
 
 Filename = root & ":\character_creator\const"
  Open Filename For Output As #14
 Print #14, con_mod
 Close #14
 
End Sub

Sub display_stats(ByRef strength As Integer, ByRef str_mod As Integer, ByRef dexterity As Integer, ByRef dex_mod As Integer, ByRef intelligence As Integer, ByRef int_mod As Integer, ByRef wisdom As Integer, ByRef wis_mod As Integer, ByRef constitution As Integer, ByRef con_mod As Integer, ByRef charisma As Integer, ByRef cha_mod As Integer, ByRef stat_block As String)

txtstats = "Stat" & vbTab & "Base" & vbTab & "Mod" & vbCrLf
txtstats = txtstats & "STR" & vbTab & strength & vbTab & str_mod & vbCrLf
txtstats = txtstats & "DEX" & vbTab & dexterity & vbTab & dex_mod & vbCrLf
txtstats = txtstats & "INT" & vbTab & intelligence & vbTab & int_mod & vbCrLf
txtstats = txtstats & "WIS" & vbTab & wisdom & vbTab & wis_mod & vbCrLf
txtstats = txtstats & "CON" & vbTab & constitution & vbTab & con_mod & vbCrLf
txtstats = txtstats & "CHA" & vbTab & charisma & vbTab & cha_mod & vbCrLf

stat_block = txtstats

End Sub

Sub naruto_occ(ByRef no_feats As Integer)
'this will not show clan occupations
Dim occupation As String

txtdisplay = "Occupation" & vbTab & vbTab & "Required Age" & vbCrLf
txtdisplay = txtdisplay & "Academy Student" & vbTab & vbTab & "10+" & vbCrLf
txtdisplay = txtdisplay & "Mentored" & vbTab & vbTab & vbTab & "12+" & vbCrLf
txtdisplay = txtdisplay & "Ninja Law Enforcement" & vbTab & "15+" & vbCrLf
txtdisplay = txtdisplay & "Ninja Technician" & vbTab & vbTab & "12+" & vbCrLf
txtdisplay = txtdisplay & "Seal Expert" & vbTab & vbTab & "15+" & vbCrLf
txtdisplay = txtdisplay & "Wandering Ninja" & vbTab & vbTab & "15+" & vbCrLf

Do
occupation = InputBox("Select your character starting occupation. Case sensitive.")
Loop Until occupation = "Academy Student" Or occupation = "Mentored" Or occupation = "Seal Expert" Or occupation = "Wandering Ninja" Or occupation = "Ninja Technician" Or occupation = "Ninja Law Enforcement"

txtmisc = txtmisc & "Occupation" & vbTab & occupation & vbCrLf

Select Case occupation
Case Is = "Academy Student"
no_feats = no_feats + 1
txtfeats = txtfeats & "Bonus feat from occupation." & vbCrLf
txtspecial = txtspecial & "Three skills from occupation." & vbCrLf

Case Is = "Mentored"
no_feats = no_feats + 1
txtfeats = txtfeats & "Bonus feat from occupation." & vbCrLf
txtspecial = txtspecial & "Three skills from occupation." & vbCrLf

Case Is = "Ninja Law Enforcement"
no_feats = no_feats + 1
txtfeats = txtfeats & "Bonus feat from occupation." & vbCrLf
txtspecial = txtspecial & "Two skills from occupation." & vbCrLf

Case Is = "Ninja Technician"
txtfeats = txtfeats & "Must have Genjutsu Adept or Ninjutsu Adept at 1st." & vbCrLf
txtspecial = txtspecial & "1 less success needed for learning lost hijutsu and kinjutsu." & vbCrLf
txtskills = txtskills & "Must have: Genjutsu 2, Knowledge (ninja lore) 2 and Ninjutsu 2." & vbCrLf

Case Is = "Seal Expert"
txtfeats = txtfeats & "Sealweaver" & vbCrLf
txtspecial = txtspecial & "Two skills from occupation." & vbCrLf

Case Is = "Wandering Ninja"
no_feats = no_feats + 1
txtfeats = txtfeats & "Bonus feat from occupation." & vbCrLf
txtspecial = txtspecial & "Two skills from occupation." & vbCrLf
End Select

End Sub

Sub modern_occ(ByRef no_feats As Integer)
Dim occupation As String

txtdisplay = "Occupation" & vbTab & vbTab & "Prerequisite" & vbCrLf
txtdisplay = txtdisplay & "Academic" & vbTab & vbTab & vbTab & "Age 23+" & vbCrLf
txtdisplay = txtdisplay & "Adventurer" & vbTab & vbTab & "Age 15+" & vbCrLf
txtdisplay = txtdisplay & "Athlete" & vbTab & vbTab & vbTab & "Str 13 or Dex 13" & vbCrLf
txtdisplay = txtdisplay & "Blue Collar" & vbTab & vbTab & "Age 18+" & vbCrLf
txtdisplay = txtdisplay & "Celebrity" & vbTab & vbTab & vbTab & "Age 15+" & vbCrLf
txtdisplay = txtdisplay & "Creative" & vbTab & vbTab & vbTab & "Age 15+" & vbCrLf
txtdisplay = txtdisplay & "Criminal" & vbTab & vbTab & vbTab & "Age 15+" & vbCrLf
txtdisplay = txtdisplay & "Dilettante" & vbTab & vbTab & vbTab & "Age 18+" & vbCrLf
txtdisplay = txtdisplay & "Doctor" & vbTab & vbTab & vbTab & "Age 25+" & vbCrLf
txtdisplay = txtdisplay & "Emergency Services" & vbTab & "Age 18+" & vbCrLf
txtdisplay = txtdisplay & "Entrepeneur" & vbTab & vbTab & "Age 18+" & vbCrLf
txtdisplay = txtdisplay & "Investigative" & vbTab & vbTab & "Age 23+" & vbCrLf
txtdisplay = txtdisplay & "Law Enforcement" & vbTab & vbTab & "Age 20+" & vbCrLf
txtdisplay = txtdisplay & "Military" & vbTab & vbTab & vbTab & "Age 18+" & vbCrLf
txtdisplay = txtdisplay & "Religious" & vbTab & vbTab & vbTab & "Age 23+" & vbCrLf
txtdisplay = txtdisplay & "Rural" & vbTab & vbTab & vbTab & "Age 15+" & vbCrLf
txtdisplay = txtdisplay & "Student" & vbTab & vbTab & vbTab & "Age 15+" & vbCrLf
txtdisplay = txtdisplay & "Technician" & vbTab & vbTab & "Age 23+" & vbCrLf
txtdisplay = txtdisplay & "White Collar" & vbTab & vbTab & "Age 23+" & vbCrLf

MsgBox "The display currently shows half of the occupations. 'UP' and 'DOWN' will change the selection."

Do
Do
occupation = InputBox("Select your character starting occupation. Case sensitive. 'UP' and 'DOWN' will change the selection")
Loop Until occupation = "UP" Or occupation = "DOWN" Or occupation = "Academic" Or occupation = "Adventurer" Or occupation = "Athlete" Or occupation = "Blue Collar" Or occupation = "Celebrity" Or occupation = "Creative" Or occupation = "Criminal" Or occupation = "Dilettante" Or occupation = "Doctor" Or occupation = "Emergency Services" Or occupation = "Entrepeneur" Or occupation = "Investigative" Or occupation = "Law Enforcement" Or occupation = "Military" Or occupation = "Religious" Or occupation = "Rural" Or occupation = "Student" Or occupation = "Technician" Or occupation = "White Collar"

Select Case occupation
Case Is = "UP"
txtdisplay.SelStart = 1 'selects the starting point of the textbox to be at the start in order for the first set of Occ's to be seen.

Case Is = "DOWN"
txtdisplay.SelStart = Len(txtdisplay.Text) 'selects the starting point of the textbox to be at the end, in order for the other set of Occ's to be seen.

Case Is = "Adventurer"
no_feats = no_feats + 1
txtfeats = txtfeats & "Bonus feat from occupation." & vbCrLf
txtspecial = txtspecial & "Two skills from occupation." & vbCrLf

Case Is = "Athlete"
no_feats = no_feats + 1
txtfeats = txtfeats & "Bonus feat from occupation." & vbCrLf
txtspecial = txtspecial & "Three skills from occupation." & vbCrLf

Case Is = "Blue Collar"
no_feats = no_feats + 1
txtfeats = txtfeats & "Bonus feat from occupation." & vbCrLf
txtspecial = txtspecial & "Three skills from occupation." & vbCrLf

Case Is = "Celebrity"
txtspecial = txtspecial & "One skill from occupation." & vbCrLf

Case Is = "Creative"
txtspecial = txtspecial & "Three skills from occupation." & vbCrLf

Case Is = "Criminal"
no_feats = no_feats + 1
txtfeats = txtfeats & "Bonus feat from occupation." & vbCrLf
txtspecial = txtspecial & "Two skills from occupation." & vbCrLf

Case Is = "Dilettante"
txtspecial = txtspecial & "One skill from occupation." & vbCrLf

Case Is = "Doctor"
txtspecial = txtspecial & "Two skills from occupation." & vbCrLf

Case Is = "Emergency Services"
txtspecial = txtspecial & "Two skills from occupation." & vbCrLf

Case Is = "Entrepreneur"
txtspecial = txtspecial & "Two skills from occupation." & vbCrLf

Case Is = "Investigative"
no_feats = no_feats + 1
txtfeats = txtfeats & "Bonus feat from occupation." & vbCrLf
txtspecial = txtspecial & "Two skills from occupation." & vbCrLf

Case Is = "Law Enforcement"
no_feats = no_feats + 1
txtfeats = txtfeats & "Bonus feat from occupation." & vbCrLf
txtspecial = txtspecial & "Two skills from occupation." & vbCrLf

Case Is = "Military"
no_feats = no_feats + 1
txtfeats = txtfeats & "Bonus feat from occupation." & vbCrLf
txtspecial = txtspecial & "Two skills from occupation." & vbCrLf

Case Is = "Religious"
txtspecial = txtspecial & "Three skills from occupation." & vbCrLf

Case Is = "Rural"
no_feats = no_feats + 1
txtfeats = txtfeats & "Bonus feat from occupation." & vbCrLf
txtspecial = txtspecial & "Two skills from occupation." & vbCrLf

Case Is = "Student"
txtspecial = txtspecial & "Three skills from occupation." & vbCrLf

Case Is = "Technician"
txtspecial = txtspecial & "Three skills from occupation." & vbCrLf

Case Is = "White Collar"
txtspecial = txtspecial & "Two skills from occupation." & vbCrLf
End Select

Loop Until occupation = "Academic" Or occupation = "Adventurer" Or occupation = "Athlete" Or occupation = "Blue Collar" Or occupation = "Celebrity" Or occupation = "Creative" Or occupation = "Criminal" Or occupation = "Dilettante" Or occupation = "Doctor" Or occupation = "Emergency Services" Or occupation = "Entrepeneur" Or occupation = "Investigative" Or occupation = "Law Enforcement" Or occupation = "Military" Or occupation = "Religious" Or occupation = "Rural" Or occupation = "Student" Or occupation = "Technician" Or occupation = "White Collar"

txtmisc = txtmisc & "Occupation" & vbTab & occupation & vbCrLf

End Sub

Private Sub cmdeditor_Click()
'automatically saves character due to testers often forgettingto save their characters
Dim counter As Integer
Dim name As String
Dim Filename As String
Dim root As String

root = InputBox$("Enter the directory that your character is saved in, i.e. D; C; X.")

Do
If Dir(root & ":", vbDirectory) = "" Then 'validates the directory
root = InputBox("Error, please enter a valid root directory.")
Else
End If
Loop Until Dir(root & ":", vbDirectory) <> ""

name = InputBox$("Enter your characters name.")

Filename = root & ":\character_creator\" & name

If Dir(Filename, vbDirectory) = "" Then
MkDir Filename
Else
End If

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
 
frmeditor.Show 'open editor
End Sub
