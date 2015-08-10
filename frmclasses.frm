VERSION 5.00
Begin VB.Form frmclasses 
   Caption         =   "classes"
   ClientHeight    =   11430
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   14310
   LinkTopic       =   "Form1"
   Picture         =   "frmclasses.frx":0000
   ScaleHeight     =   11430
   ScaleWidth      =   14310
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdget_info 
      Caption         =   "Get class data"
      Height          =   495
      Left            =   8400
      Picture         =   "frmclasses.frx":8DB1D
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox txtstats 
      Height          =   2175
      Left            =   8160
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   6840
      Width           =   1935
   End
   Begin VB.TextBox txtspells 
      Height          =   1095
      Left            =   9960
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   2640
      Width           =   3975
   End
   Begin VB.TextBox txtspecial 
      Height          =   7575
      Left            =   10200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   4200
      Width           =   3735
   End
   Begin VB.TextBox txtclasses 
      Height          =   3375
      Left            =   8280
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txtmods 
      Height          =   2175
      Left            =   8400
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   9480
      Width           =   1455
   End
   Begin VB.TextBox txtdisplay 
      Height          =   1815
      Left            =   9960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   360
      Width           =   3975
   End
   Begin VB.CommandButton cmdgetclass 
      Caption         =   "Choose Classes"
      Height          =   615
      Left            =   8280
      Picture         =   "frmclasses.frx":2488CF
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Go back to Character Creator"
      Height          =   735
      Left            =   8400
      Picture         =   "frmclasses.frx":403681
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox txtprestclasses 
      Height          =   11295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   360
      Width           =   7935
   End
   Begin VB.CommandButton cmdend 
      Caption         =   "End Program"
      Height          =   375
      Left            =   8520
      Picture         =   "frmclasses.frx":5BE433
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Information on choosing classes"
      Height          =   255
      Left            =   10680
      TabIndex        =   16
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "Special Abilities"
      Height          =   255
      Left            =   11400
      TabIndex        =   15
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Combat Stats"
      Height          =   255
      Left            =   8640
      TabIndex        =   14
      Top             =   9240
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Stats/Abilities"
      Height          =   255
      Left            =   8640
      TabIndex        =   13
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Spells/Day"
      Height          =   255
      Left            =   11400
      TabIndex        =   12
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Base Classes"
      Height          =   255
      Left            =   8520
      TabIndex        =   11
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Prestige Classes"
      Height          =   255
      Left            =   3240
      TabIndex        =   10
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmclasses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdback_Click()
frmcharsheet.Show
Unload frmclasses
End Sub

Private Sub cmdend_Click()
End
End Sub

Private Sub cmdgetclass_Click()
Dim ba As Integer
Dim will_save As Integer
Dim ref_save As Integer
Dim fort_save As Integer
Dim hp As Integer
Dim skill_pts As Integer
Dim ac As Integer
Dim feats As Integer

Dim b_att(3, 20) As Integer '1 is good, 2 is med, 3 is bad
Dim will(3, 20) As Integer
Dim ref(3, 20) As Integer
Dim fort(3, 20) As Integer

Dim int_mod As Integer
Dim con_mod As Integer
Dim dex_mod As Integer
Dim str_mod As Integer
Dim wis_mod As Integer
Dim cha_mod As Integer

Dim system As String
Dim level As Integer

Dim counter As Integer

Dim skills As Integer
Dim feat As Integer
Dim strength As Integer
Dim charisma As Integer
Dim intelligence As Integer
Dim constitution As Integer
Dim dexterity As Integer
Dim wisdom As Integer

'populates save arrays
b_att(2, 1) = 0
b_att(3, 1) = 0
will(1, 1) = 2
ref(1, 1) = 2
fort(1, 1) = 2
will(3, 1) = 0
will(3, 2) = 0
ref(3, 1) = 0
ref(3, 2) = 0
fort(3, 1) = 0
fort(3, 2) = 0


For counter = 1 To 20
b_att(1, counter) = counter
Next counter

counter = 0

For counter = 1 To 5
b_att(2, counter * 4 - 2) = counter * 3 - 2
b_att(2, counter * 4 - 1) = counter * 3 - 1
b_att(2, counter * 4) = counter * 3
If counter <= 4 Then
b_att(2, counter * 4 + 1) = counter * 3
End If
Next counter

counter = 0

For counter = 1 To 10
b_att(3, 2 * counter) = counter
If counter <= 9 Then
b_att(3, 2 * counter + 1) = b_att(3, 2 * counter)
End If
Next counter

counter = 0

For counter = 1 To 10
will(1, 2 * counter) = counter + 2
If counter <= 9 Then
will(1, 2 * counter + 1) = will(1, 2 * counter)
End If
Next counter

counter = 0

For counter = 1 To 10
ref(1, 2 * counter) = counter + 2
If counter <= 9 Then
ref(1, 2 * counter + 1) = ref(1, 2 * counter)
End If
Next counter

counter = 0

For counter = 1 To 10
fort(1, 2 * counter) = counter + 2
If counter <= 9 Then
fort(1, 2 * counter + 1) = fort(1, 2 * counter)
End If
Next counter

counter = 0

For counter = 1 To 6
will(3, 3 * counter) = counter
will(3, 3 * counter + 1) = will(3, 3 * counter)
will(3, 3 * counter + 2) = will(3, 3 * counter)
Next counter

counter = 0

For counter = 1 To 6
ref(3, 3 * counter) = counter
ref(3, 3 * counter + 1) = ref(3, 3 * counter)
ref(3, 3 * counter + 2) = ref(3, 3 * counter)
Next counter

counter = 0

For counter = 1 To 6
fort(3, 3 * counter) = counter
fort(3, 3 * counter + 1) = fort(3, 3 * counter)
fort(3, 3 * counter + 2) = fort(3, 3 * counter)
Next counter

'these arrays are given individually, as to calculate them with a loop takes up more code than to list them seperately

will(2, 1) = 1
will(2, 2) = 2
will(2, 3) = 2
will(2, 4) = 2
will(2, 5) = 3
will(2, 6) = 3
will(2, 7) = 4
will(2, 8) = 4
will(2, 9) = 4
will(2, 10) = 5

ref(2, 1) = 1
ref(2, 2) = 2
ref(2, 3) = 2
ref(2, 4) = 2
ref(2, 5) = 3
ref(2, 6) = 3
ref(2, 7) = 4
ref(2, 8) = 4
ref(2, 9) = 4
ref(2, 10) = 5

fort(2, 1) = 1
fort(2, 2) = 2
fort(2, 3) = 2
fort(2, 4) = 2
fort(2, 5) = 3
fort(2, 6) = 3
fort(2, 7) = 4
fort(2, 8) = 4
fort(2, 9) = 4
fort(2, 10) = 5

root = InputBox("Insert the letter where you made your save folder.")

Do
If Dir(root & ":", vbDirectory) = "" Then
root = InputBox("Error, please enter a valid root directory.")
Else
End If
Loop Until Dir(root & ":", vbDirectory) <> ""

 Filename = root & ":\character_creator\system"
 Open Filename For Input As #1
 system = Input$(LOF(1), 1)
Close #1

Filename = root & ":\character_creator\level"
 Open Filename For Input As #2
 level = Input$(LOF(2), 2)
Close #2

Filename = root & ":\character_creator\dex"
Open Filename For Input As #3
dex_mod = Input$(LOF(3), 3)
Close #3

Filename = root & ":\character_creator\int"
Open Filename For Input As #4
int_mod = Input$(LOF(4), 4)
Close #4

Filename = root & ":\character_creator\const"
Open Filename For Input As #5
con_mod = Input$(LOF(5), 5)
Close #5

Filename = root & ":\character_creator\feat"
Open Filename For Input As #6
feats = Input$(LOF(6), 6)
Close #6

Filename = root & ":\character_creator\skill"
Open Filename For Input As #7
skills = Input(LOF(7), 7)
Close #7

Filename = root & ":\character_creator\strength"
Open Filename For Input As #8
strength = Input(LOF(8), 8)
Close #8

Filename = root & ":\character_creator\charisma"
Open Filename For Input As #9
charisma = Input(LOF(9), 9)
Close #9

Filename = root & ":\character_creator\intelligence"
Open Filename For Input As #10
intelligence = Input(LOF(10), 10)
Close #10

Filename = root & ":\character_creator\constitution"
Open Filename For Input As #11
constitution = Input(LOF(11), 11)
Close #11

Filename = root & ":\character_creator\dexterity"
Open Filename For Input As #12
dexterity = Input(LOF(12), 12)
Close #12

Filename = root & ":\character_creator\wisdom"
Open Filename For Input As #13
wisdom = Input(LOF(13), 13)
Close #13

If level <= 0 Then
MsgBox ("Due to Level Adjustment, your level has fallen below 1. You have 1 level to spend, but will have to pay off your LA at a later date.")
level = 1
End If

MsgBox ("You are currently level " & level & ", and are using the " & system & " system.")

Select Case system
Case Is = "D&D" & vbCrLf
txtdisplay = "You are using the D&D 3.5 rules and system. A list of available classes are opposite. Please note that it is impossible to prestige before level 4, so a character must have at least 3 levels in a basic class in order to meet any of the prerequisites. The prereq's for each class (if any) are indented under their names."
Call dnd_classes(level, ba, b_att(), will_save, will(), ref_save, ref(), fort_save, fort(), hp, con_mod, skill_pts, int_mod, ac, feats, strength, charisma, intelligence, constitution)

Case Is = "D20 Modern" & vbCrLf
txtdisplay = "You are using the D20 Modern rules and system. A list of available classes are opposite. Please note that it is impossible to prestige before level 4, so a character must have at least 3 levels in a basic class in order to meet any of the prerequisites. The prereq's for each class (if any) are indented under their names."
Call mod_classes(level, ba, b_att(), will_save, will(), ref_save, ref(), fort_save, fort(), hp, con_mod, skill_pts, int_mod, ac, feats)

Case Is = "Naruto D20" & vbCrLf
txtdisplay = "You are using the Naruto D20 rules and system. A list of available classes are opposite. Please note that it is impossible to prestige before level 4, so a character must have at least 3 levels in a basic class in order to meet any of the prerequisites. The prereq's for each class (if any) are indented under their names."
Call naruto_classes(level, ba, b_att(), will_save, will(), ref_save, ref(), fort_save, fort(), hp, con_mod, skill_pts, int_mod, ac, feats)
End Select

str_mod = ((strength - 10) / 2) - 0.25
dex_mod = ((dexterity - 10) / 2) - 0.25
int_mod = ((intelligence - 10) / 2) - 0.25
wis_mod = ((wisdom - 10) / 2) - 0.25
con_mod = ((constitution - 10) / 2) - 0.25
cha_mod = ((charisma - 10) / 2) - 0.25

txtstats = "Stat" & vbTab & "Base" & vbTab & "Mod" & vbCrLf
txtstats = txtstats & "STR" & vbTab & strength & vbTab & str_mod & vbCrLf
txtstats = txtstats & "DEX" & vbTab & dexterity & vbTab & dex_mod & vbCrLf
txtstats = txtstats & "INT" & vbTab & intelligence & vbTab & int_mod & vbCrLf
txtstats = txtstats & "WIS" & vbTab & wisdom & vbTab & wis_mod & vbCrLf
txtstats = txtstats & "CON" & vbTab & constitution & vbTab & con_mod & vbCrLf
txtstats = txtstats & "CHA" & vbTab & charisma & vbTab & cha_mod & vbCrLf

stat_block = txtstats

frmcharsheet.txtstats.Text = txtstats.Text

skill_pts = skill_pts + skills
ac = ac + dex_mod + 10
frmcharsheet.txtskills.Text = "You have " & skill_pts & " skill points to spend. Class skills can be found in the rulebooks."
frmcharsheet.txtfeats.Text = frmcharsheet.txtfeats.Text & "You have " & Left(feats, 2) & " feat(s) to choose, including bonus feats. Bonus feats from classes are shown in the specials column, and should be chosen from the appropiate list."

If system = "Naruto D20" & vbCrLf Then
Call calc_chakra(level, con_mod)
End If

txtmods.Text = txtmods.Text & vbCrLf & "AC/Def" & vbTab & ac
frmcharsheet.txtmods.Text = txtmods.Text
frmcharsheet.txtspecial.Text = frmcharsheet.txtspecial.Text & txtspecial.Text
frmcharsheet.txtspells.Text = txtspells.Text
End Sub

Sub calc_chakra(ByVal level As Integer, ByVal con_mod As Integer)
Dim chakra_pool As Integer
Dim chakra_res As Integer

chakra_pool = 2 * (level + 1) + (con_mod * lvl)
chakra_res = 2 * level

txtmods.Text = txtmods.Text & vbCrLf & "Chakra" & vbTab & chakra_pool & vbCrLf & "Chakra R" & vbTab & chakra_res
End Sub

Sub dnd_classes(ByVal level As Integer, ByRef ba As Integer, ByRef b_att() As Integer, ByRef will_save As Integer, ByRef will() As Integer, ByRef ref_save As Integer, ByRef ref() As Integer, ByRef fort_save As Integer, ByRef fort() As Integer, ByRef hp As Integer, ByVal con_mod As Integer, ByRef skill_pts As Integer, ByVal int_mod As Integer, ByRef ac As Integer, ByRef feats As Integer, ByRef strength As Integer, ByRef charisma As Integer, ByRef intelligence As Integer, ByRef constitution As Integer)

Dim counter(3) As Integer
Dim Class(26) As Integer 'finds out the levels in each class
Dim lvl As Integer 'individual level in class
Dim ttl_lvl As Integer 'total levels used
Dim cls As String
Dim chosen(26) As String
Dim selected As Integer
Dim select_cls As String
Dim increment As Integer
Dim first_class As String

ttl_lvl = 0
counter(2) = 1
counter(3) = 1

Do
For counter(1) = 1 To level
If ttl_lvl < level Then

If ttl_lvl < 3 Then 'makes sure that the first 3 levels are not prestige
Do
cls = InputBox("select the class that you wish to take. Initial letter must be uppercase, i.e. Barbarian, Bard. Can not be prestige.")
Loop Until cls = "Barbarian" Or cls = "Bard" Or cls = "Cleric" Or cls = "Druid" Or cls = "Fighter" Or cls = "Monk" Or cls = "Paladin" Or cls = "Ranger" Or cls = "Rogue" Or cls = "Sorcerer" Or cls = "Wizard"

Else 'if you have taken more than 3 levels, you can then prestige
Do
cls = InputBox("Select the class that you wish to take. Initial letter must be uppercase, i.e. Barbarian, Bard, Assassin.")
Loop Until cls = "Barbarian" Or cls = "Bard" Or cls = "Cleric" Or cls = "Druid" Or cls = "Fighter" Or cls = "Monk" Or cls = "Paladin" Or cls = "Ranger" Or cls = "Rogue" Or cls = "Sorcerer" Or cls = "Wizard" Or cls = "Arcane Archer" Or cls = "Arcane Trickster" Or cls = "Archmage" Or cls = "Assassin" Or cls = "Blackguard" Or cls = "Dragon Disciple" Or cls = "Duelist" Or cls = "Dwarven Defender" Or cls = "Eldritch Knight" Or cls = "Hierophant" Or cls = "Horizon Walker" Or cls = "Loremaster" Or cls = "Mystic Theurge" Or cls = "Shadowdancer" Or cls = "Thaumaturgist"
End If

If counter(1) = 1 Then
first_class = cls
End If

chosen(counter(2)) = cls

Do
lvl = InputBox("Select how many levels you wish to take in that class. You have " & level - ttl_lvl & " levels left to spend.")

'makes sure that you can't take over the maximum amount of levels in a class

If cls = "Barbarian" Or cls = "Bard" Or cls = "Cleric" Or cls = "Druid" Or cls = "Fighter" Or cls = "Monk" Or cls = "Paladin" Or cls = "Ranger" Or cls = "Rogue" Or cls = "Sorcerer" Or cls = "Wizard" Then

Do
If lvl > 20 Or lvl < 0 Then
lvl = InputBox("Max of 20 levels in class you wish to take in that class.")
End If
Loop Until lvl <= 20
End If


If cls = "Arcane Archer" Or cls = "Arcane Trickster" Or cls = "Assassin" Or cls = "Blackguard" Or cls = "Dragon Disciple" Or cls = "Duelist" Or cls = "Dwarven Defender" Or cls = "Eldritch Knight" Or cls = "Horizon Walker" Or cls = "Loremaster" Or cls = "Mystic Theurge" Or cls = "Shadowdancer" Then

Do

If lvl > 10 Or lvl < 0 Then
lvl = InputBox("Max of 10 levels in class you wish to take in that class.")
End If
Loop Until lvl <= 10
End If


If cls = "Archmage" Or cls = "Hierophant" Or cls = "Thaumaturgist" Then

Do

If lvl > 5 Or lvl < 0 Then
lvl = InputBox("Max of 5 levels in class you wish to take in that class.")
End If
Loop Until lvl <= 5
End If


Loop Until lvl <= level 'the user can't continue if he has more than his maximum level
ttl_lvl = ttl_lvl + lvl



Select Case cls 'finds out how many levels are in each class, and resets the choice if you end up with more levels than you should, i.e. take archmage/5 twice
Case Is = "Barbarian"
Class(1) = Class(1) + lvl
If Class(1) > 20 Then
Class(1) = Class(1) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Bard"
Class(2) = Class(2) + lvl
If Class(2) > 20 Then
Class(2) = Class(2) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Cleric"
Class(3) = Class(3) + lvl
If Class(3) > 20 Then
Class(3) = Class(3) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Druid"
Class(4) = Class(4) + lvl
If Class(4) > 20 Then
Class(4) = Class(4) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Fighter"
Class(5) = Class(5) + lvl
If Class(5) > 20 Then
Class(5) = Class(5) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Monk"
Class(6) = Class(6) + lvl
If Class(6) > 20 Then
Class(6) = Class(6) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Paladin"
Class(7) = Class(7) + lvl
If Class(7) > 20 Then
Class(7) = Class(7) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Ranger"
Class(8) = Class(8) + lvl
If Class(8) > 20 Then
Class(8) = Class(8) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Rogue"
Class(9) = Class(9) + lvl
If Class(9) > 20 Then
Class(9) = Class(9) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Sorcerer"
Class(10) = Class(10) + lvl
If Class(10) > 20 Then
Class(10) = Class(1) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Wizard"
Class(11) = Class(11) + lvl
If Class(11) > 20 Then
Class(11) = Class(11) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Arcane Archer"
Class(12) = Class(12) + lvl
If Class(12) > 10 Then
Class(12) = Class(12) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Arcane Trickster"
Class(13) = Class(13) + lvl
If Class(13) > 10 Then
Class(13) = Class(13) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Archmage"
Class(14) = Class(14) + lvl
If Class(14) > 5 Then
Class(14) = Class(14) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Assassin"
Class(15) = Class(15) + lvl
If Class(15) > 10 Then
Class(15) = Class(15) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Blackguard"
Class(16) = Class(16) + lvl
If Class(16) > 10 Then
Class(16) = Class(16) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Dragon Disciple"
Class(17) = Class(17) + lvl
If Class(17) > 10 Then
Class(17) = Class(17) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Duelist"
Class(18) = Class(18) + lvl
If Class(18) > 10 Then
Class(18) = Class(18) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Dwarven Defender"
Class(19) = Class(19) + lvl
If Class(19) > 10 Then
Class(19) = Class(19) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Eldritch Knight"
Class(20) = Class(20) + lvl
If Class(20) > 10 Then
Class(20) = Class(20) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Hierophant"
Class(21) = Class(21) + lvl
If Class(21) > 5 Then
Class(21) = Class(1) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Horizon Walker"
Class(22) = Class(22) + lvl
If Class(22) > 10 Then
Class(22) = Class(22) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Loremaster"
Class(23) = Class(23) + lvl
If Class(23) > 10 Then
Class(23) = Class(23) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Mystic Theurge"
Class(24) = Class(24) + lvl
If Class(24) > 10 Then
Class(24) = Class(24) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Shadowdancer"
Class(25) = Class(25) + lvl
If Class(25) > 10 Then
Class(25) = Class(25) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Thaumaturgist"
Class(26) = Class(26) + lvl
If Class(26) > 5 Then
Class(26) = Class(26) - lvl
MsgBox "Class can not have that many levels in it."
End If

End Select

counter(2) = counter(2) + 1

End If

Next counter(1)

Loop Until ttl_lvl = level


increment = 1

Do 'makes sure that no duplicates appear in chosen class, so that if fighter is chosen for 3 levels, then 2 levels, it won't give fighter 5 + fighter 5

select_cls = chosen(increment)

For increment = 1 To counter(2)
selected = increment + 1
If select_cls = chosen(selected) Then
chosen(selected) = "x"
End If
Next increment

Dim spec_loop As Integer 'gets the specials from classes
Dim clslvls As String 'puts down how many levels are in each class, which will be put into the misc area afterwards

increment = 1

For increment = 1 To counter(2)
select_cls = chosen(increment)

Select Case select_cls

Case Is = "Barbarian"
ba = ba + b_att(1, Class(1))
will_save = will_save + will(3, Class(1))
ref_save = ref_save + ref(3, Class(1))
fort_save = fort_save + fort(1, Class(1))
hp = hp + (12 + con_mod) * Class(1)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(1)

If first_class = select_cls Then
skill_pts = skill_pts + (4 + int_mod) * (Class(1) + 3)
Else
skill_pts = skill_pts + (4 + int_mod) * Class(1)
End If

Select Case Class(1) 'gets rage/day
Case Is = 20
txtspecial = txtspecial & "Rage 6/Day" & vbCrLf
Case Is >= 16
txtspecial = txtspecial & "Rage 5/Day" & vbCrLf
Case Is >= 12
txtspecial = txtspecial & "Rage 4/Day" & vbCrLf
Case Is >= 8
txtspecial = txtspecial & "Rage 3/Day" & vbCrLf
Case Is >= 4
txtspecial = txtspecial & "Rage 2/Day" & vbCrLf
Case Is >= 1
txtspecial = txtspecial & "Rage 1/Day" & vbCrLf
End Select

Select Case Class(1) 'gets Damage Reduction/-
Case Is >= 19
txtspecial = txtspecial & "Damage Reduction 5/-" & vbCrLf
Case Is >= 16
txtspecial = txtspecial & "Damage Reduction 4/-" & vbCrLf
Case Is >= 13
txtspecial = txtspecial & "Damage Reduction 3/-" & vbCrLf
Case Is >= 10
txtspecial = txtspecial & "Damage Reduction 2/-" & vbCrLf
Case Is >= 7
txtspecial = txtspecial & "Damage Reduction 1/-" & vbCrLf
End Select

Select Case Class(1) 'gets trap sense
Case Is >= 18
txtspecial = txtspecial & "Trap Sense +6" & vbCrLf
Case Is >= 15
txtspecial = txtspecial & "Trap Sense +5" & vbCrLf
Case Is >= 12
txtspecial = txtspecial & "Trap Sense +4" & vbCrLf
Case Is >= 9
txtspecial = txtspecial & "Trap Sense +3" & vbCrLf
Case Is >= 6
txtspecial = txtspecial & "Trap Sense +2" & vbCrLf
Case Is >= 3
txtspecial = txtspecial & "Trap Sense +1" & vbCrLf
End Select

spec_loop = Class(1)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 20
txtspecial = txtspecial & "Mighty Rage" & vbCrLf
Case Is = 17
txtspecial = txtspecial & "Tireless Rage" & vbCrLf
Case Is = 14
txtspecial = txtspecial & "Indomitable Will" & vbCrLf
Case Is = 11
txtspecial = txtspecial & "Greater Rage" & vbCrLf
Case Is = 5
txtspecial = txtspecial & "Improved Uncanny Dodge" & vbCrLf
Case Is = 2
txtspecial = txtspecial & "Uncanny Dodge" & vbCrLf
Case Is = 1
txtspecial = txtspecial & "Fast Movement" & vbCrLf & "Illiteracy" & vbCrLf
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Bard"
ba = ba + b_att(2, Class(2))
will_save = will_save + will(1, Class(2))
ref_save = ref_save + ref(1, Class(2))
fort_save = fort_save + fort(3, Class(2))
hp = hp + (6 + con_mod) * Class(2)

clslvls = clslvls & Left(select_cls, 4) & "/" & Class(2)

If first_class = select_cls Then
skill_pts = skill_pts + (6 + int_mod) * (Class(2) + 3)
Else
skill_pts = skill_pts + (6 + int_mod) * Class(2)
End If

Select Case Class(2) 'gets inspire courage
Case Is = 20
txtspecial = txtspecial & "Inspire Courage +4" & vbCrLf
Case Is >= 14
txtspecial = txtspecial & "Inspire Courage +3" & vbCrLf
Case Is >= 8
txtspecial = txtspecial & "Inspire Courage +2" & vbCrLf
Case Is >= 1
txtspecial = txtspecial & "Inspire Courage +1" & vbCrLf
End Select

spec_loop = Class(2)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 18
txtspecial = txtspecial & "Mass Suggestion" & vbCrLf
Case Is = 15
txtspecial = txtspecial & "Inspire heroics" & vbCrLf
Case Is = 12
txtspecial = txtspecial & "Song of Freedom" & vbCrLf & "" & vbCrLf
Case Is = 9
txtspecial = txtspecial & "Inspire Greatness" & vbCrLf
Case Is = 6
txtspecial = txtspecial & "Suggestion" & vbCrLf
Case Is = 3
txtspecial = txtspecial & "Inspire Competance" & vbCrLf
Case Is = 1
txtspecial = txtspecial & "Bardic Music" & vbCrLf & "Bardic Knowledge" & vbCrLf & "Countersong" & vbCrLf & "Fascinate" & vbCrLf
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0

If Class(2) >= 1 Then
txtspells = txtspells & "Bardic Spells/Day" & vbCrLf
End If

Select Case Class(2) 'gets spells/day
Case Is = 20
txtspells = txtspells & "0:4  1:4  2:4  3:4  4:4  5:4  6:4" & vbCrLf
Case Is = 19
txtspells = txtspells & "0:4  1:4  2:4  3:4  4:4  5:4  6:3" & vbCrLf
Case Is = 18
txtspells = txtspells & "0:4  1:4  2:4  3:4  4:4  5:3  6:2" & vbCrLf
Case Is = 17
txtspells = txtspells & "0:4  1:4  2:4  3:4  4:3  5:3  6:1" & vbCrLf
Case Is = 16
txtspells = txtspells & "0:4  1:4  2:4  3:3  4:3  5:2  6:0" & vbCrLf
Case Is = 15
txtspells = txtspells & "0:4  1:4  2:3  3:3  4:3  5:2" & vbCrLf
Case Is = 14
txtspells = txtspells & "0:4  1:3  2:3  3:3  4:3  5:1" & vbCrLf
Case Is = 13
txtspells = txtspells & "0:3  1:3  2:3  3:3  4:2  5:0" & vbCrLf
Case Is = 12
txtspells = txtspells & "0:3  1:3  2:3  3:3  4:2" & vbCrLf
Case Is = 11
txtspells = txtspells & "0:3  1:3  2:3  3:3  4:1" & vbCrLf
Case Is = 10
txtspells = txtspells & "0:3  1:3  2:3  3:2  4:0" & vbCrLf
Case Is = 9
txtspells = txtspells & "0:3  1:3  2:3  3:2" & vbCrLf
Case Is = 8
txtspells = txtspells & "0:3  1:3  2:3  3:1" & vbCrLf
Case Is = 7
txtspells = txtspells & "0:3  1:3  2:2  3:0" & vbCrLf
Case Is = 6
txtspells = txtspells & "0:3  1:3  2:2" & vbCrLf
Case Is = 5
txtspells = txtspells & "0:3  1:3  2:1" & vbCrLf
Case Is = 4
txtspells = txtspells & "0:3  1:2  2:0" & vbCrLf
Case Is = 3
txtspells = txtspells & "0:3  1:1" & vbCrLf
Case Is = 2
txtspells = txtspells & "0:3  1:0" & vbCrLf
Case Is = 1
txtspells = txtspells & "0:2" & vbCrLf
End Select


Case Is = "Cleric"
ba = ba + b_att(2, Class(3))
will_save = will_save + will(1, Class(3))
ref_save = ref_save + ref(3, Class(3))
fort_save = fort_save + fort(1, Class(3))
hp = hp + (8 + con_mod) * Class(3)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(3)

If first_class = select_cls Then
skill_pts = skill_pts + (2 + int_mod) * (Class(3) + 3)
Else
skill_pts = skill_pts + (2 + int_mod) * Class(3)
End If

If Class(3) >= 1 Then
txtspecial = txtspecial & "Turn or Rebuke Undead" & vbCrLf
txtspells = txtspells & "Clerical Spells/Day - Divine" & vbCrLf
End If

Select Case Class(3) 'gets spells/day
Case Is = 20
txtspells = txtspells & "0:6  1:5+1  2:5+1  3:5+1  4:5+1  5:5+1  6:4+1  7:4+1  8:4+1  9:4+1" & vbCrLf
Case Is = 19
txtspells = txtspells & "0:6  1:5+1  2:5+1  3:5+1  4:5+1  5:5+1  6:4+1  7:4+1  8:3+1  9:3+1" & vbCrLf
Case Is = 18
txtspells = txtspells & "0:6  1:5+1  2:5+1  3:5+1  4:5+1  5:4+1  6:4+1  7:3+1  8:3+1  9:2+1" & vbCrLf
Case Is = 17
txtspells = txtspells & "0:6  1:5+1  2:5+1  3:5+1  4:5+1  5:4+1  6:4+1  7:3+1  8:2+1  9:1+1" & vbCrLf
Case Is = 16
txtspells = txtspells & "0:6  1:5+1  2:5+1  3:5+1  4:4+1  5:4+1  6:3+1  7:3+1  8:2+1" & vbCrLf
Case Is = 15
txtspells = txtspells & "0:6  1:5+1  2:5+1  3:5+1  4:4+1  5:4+1  6:3+1  7:2+1  8:1+1" & vbCrLf
Case Is = 14
txtspells = txtspells & "0:6  1:5+1  2:5+1  3:4+1  4:4+1  5:3+1  6:3+1  7:2+1" & vbCrLf
Case Is = 13
txtspells = txtspells & "0:6  1:5+1  2:5+1  3:4+1  4:4+1  5:3+1  6:2+1  7:1+1" & vbCrLf
Case Is = 12
txtspells = txtspells & "0:6  1:5+1  2:4+1  3:4+1  4:3+1  5:3+1  6:2+1" & vbCrLf
Case Is = 11
txtspells = txtspells & "0:6  1:5+1  2:4+1  3:4+1  4:3+1  5:2+1  6:1+1" & vbCrLf
Case Is = 10
txtspells = txtspells & "0:6  1:5+1  2:4+1  3:3+1  4:3+1  5:2+1" & vbCrLf
Case Is = 9
txtspells = txtspells & "0:6  1:4+1  2:4+1  3:3+1  4:2+1  5:1+1" & vbCrLf
Case Is = 8
txtspells = txtspells & "0:6  1:4+1  2:3+1  3:3+1  4:2+1" & vbCrLf
Case Is = 7
txtspells = txtspells & "0:6  1:4+1  2:3+1  3:2+1  4:1+1" & vbCrLf
Case Is = 6
txtspells = txtspells & "0:5  1:3+1  2:3+1  3:2+1" & vbCrLf
Case Is = 5
txtspells = txtspells & "0:5  1:3+1  2:2+1  3:1+1" & vbCrLf
Case Is = 4
txtspells = txtspells & "0:5  1:3+1  2:2+1" & vbCrLf
Case Is = 3
txtspells = txtspells & "0:4  1:2+1  2:1+1" & vbCrLf
Case Is = 2
txtspells = txtspells & "0:4  1:2+1" & vbCrLf
Case Is = 1
txtspells = txtspells & "0:3  1:1+1" & vbCrLf
End Select


Case Is = "Druid"
ba = ba + b_att(2, Class(4))
will_save = will_save + will(1, Class(4))
ref_save = ref_save + ref(3, Class(4))
fort_save = fort_save + fort(1, Class(4))
hp = hp + (8 + con_mod) * Class(4)

clslvls = clslvls & Left(select_cls, 4) & "/" & Class(4)

If first_class = select_cls Then
skill_pts = skill_pts + (4 = int_mod) * (Class(4) + 3)
Else
skill_pts = skill_pts + (4 + int_mod) * Class(4)
End If

Select Case Class(4) 'gets wild shape/day
Case Is >= 18
txtspecial = txtspecial & "Wild Shape 6/Day" & vbCrLf
Case Is >= 14
txtspecial = txtspecial & "Wild Shape 5/Day" & vbCrLf
Case Is >= 10
txtspecial = txtspecial & "Wild Shape 4/Day" & vbCrLf
Case Is >= 7
txtspecial = txtspecial & "Wild Shape 3/Day" & vbCrLf
Case Is >= 6
txtspecial = txtspecial & "Wild Shape 2/Day" & vbCrLf
Case Is >= 5
txtspecial = txtspecial & "Wild Shape 1/Day" & vbCrLf
End Select

Select Case Class(4) 'gets wild shape/day
Case Is = 20
txtspecial = txtspecial & "Wild Shape(Elemental) 3/Day" & vbCrLf
Case Is >= 18
txtspecial = txtspecial & "Wild Shape(Elemental) 2/Day" & vbCrLf
Case Is >= 16
txtspecial = txtspecial & "Wild Shape(Elemental) 1/Day" & vbCrLf
End Select

spec_loop = Class(4)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 20
txtspecial = txtspecial & "Wild Shape (Huge Elemental)" & vbCrLf
Case Is = 15
txtspecial = txtspecial & "Timeless Body" & vbCrLf & "Wild Shape (Huge)" & vbCrLf
Case Is = 13
txtspecial = txtspecial & "A Thousand Faces" & vbCrLf
Case Is = 12
txtspecial = txtspecial & "Wild Shape (Plant)" & vbCrLf
Case Is = 11
txtspecial = txtspecial & "Wild Shape (Tiny)" & vbCrLf
Case Is = 9
txtspecial = txtspecial & "Venom Immunity" & vbCrLf
Case Is = 8
txtspecial = txtspecial & "Wild Shape (Large)" & vbCrLf
Case Is = 4
txtspecial = txtspecial & "Resist Nature's Lure" & vbCrLf
Case Is = 3
txtspecial = txtspecial & "Trackless Step" & vbCrLf
Case Is = 2
txtspecial = txtspecial & "Woodland Stride" & vbCrLf
Case Is = 1
txtspecial = txtspecial & "Animal Companion" & vbCrLf & "Nature Sense" & vbCrLf & "Wild Empathy" & vbCrLf
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0

If Class(4) >= 1 Then
txtspells = txtspells & "Druidic Spells/Day" & vbCrLf
End If

Select Case Class(4) 'gets spells/day
Case Is = 20
txtspells = txtspells & "0:6  1:5  2:5  3:5  4:5  5:5  6:4  7:4  8:4  9:4" & vbCrLf
Case Is = 19
txtspells = txtspells & "0:6  1:5  2:5  3:5  4:5  5:5  6:4  7:4  8:3  9:3" & vbCrLf
Case Is = 18
txtspells = txtspells & "0:6  1:5  2:5  3:5  4:5  5:4  6:4  7:3  8:3  9:2 " & vbCrLf
Case Is = 17
txtspells = txtspells & "0:6  1:5  2:5  3:5  4:5  5:4  6:4  7:3  8:2  9:1 " & vbCrLf
Case Is = 16
txtspells = txtspells & "0:6  1:5  2:5  3:5  4:4  5:4  6:3  7:3  8:2 " & vbCrLf
Case Is = 15
txtspells = txtspells & "0:6  1:5  2:5  3:5  4:4  5:4  6:3  7:2  8:1 " & vbCrLf
Case Is = 14
txtspells = txtspells & "0:6  1:5  2:5  3:4  4:4  5:3  6:3  7:2 " & vbCrLf
Case Is = 13
txtspells = txtspells & "0:6  1:5  2:5  3:4  4:4  5:3  6:2  7:1 " & vbCrLf
Case Is = 12
txtspells = txtspells & "0:6  1:5  2:4  3:4  4:3  5:3  6:2 " & vbCrLf
Case Is = 11
txtspells = txtspells & "0:6  1:5  2:4  3:4  4:3  5:2  6:1 " & vbCrLf
Case Is = 10
txtspells = txtspells & "0:6  1:5  2:4  3:3  4:3  5:2 " & vbCrLf
Case Is = 9
txtspells = txtspells & "0:6  1:4  2:4  3:3  4:2  5:1 " & vbCrLf
Case Is = 8
txtspells = txtspells & "0:6  1:4  2:3  3:3  4:2 " & vbCrLf
Case Is = 7
txtspells = txtspells & "0:6  1:4  2:3  3:2  4:1 " & vbCrLf
Case Is = 6
txtspells = txtspells & "0:5  1:3  2:3  3:2 " & vbCrLf
Case Is = 5
txtspells = txtspells & "0:5  1:3  2:2  3:1 " & vbCrLf
Case Is = 4
txtspells = txtspells & "0:5  1:3  2:2 " & vbCrLf
Case Is = 3
txtspells = txtspells & "0:4  1:2  2:1 " & vbCrLf
Case Is = 2
txtspells = txtspells & "0:4  1:2 " & vbCrLf
Case Is = 1
txtspells = txtspells & "0:3  1:1 " & vbCrLf
End Select


Case Is = "Fighter"
ba = ba + b_att(1, Class(5))
will_save = will_save + will(3, Class(5))
ref_save = ref_save + ref(3, Class(5))
fort_save = fort_save + fort(1, Class(5))
hp = hp + (10 + con_mod) * Class(5)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(5)

If first_class = select_cls Then
skill_pts = skill_pts + (2 + int_mod) * (Class(5) + 3)
Else
skill_pts = skill_pts + (2 + int_mod) * Class(5)
End If

If Class(5) >= 1 Then
feats = feats + 1
End If

If Class(5) <> 2 And Class(5) <> 4 And Class(5) <> 6 And Class(5) <> 8 And Class(5) <> 10 And Class(5) <> 12 And Class(5) <> 14 And Class(5) <> 16 And Class(5) <> 18 And Class(5) <> 20 Then
feats = feats + ((Class(5) - 1) / 2)
txtspecial = txtspecial & ((Class(5) - 1) / 2) + 1 & " Fighter Bonus Feats" & vbCrLf
Else
feats = feats + Class(5) / 2
txtspecial = txtspecial & (Class(5) / 2) + 1 & " Fighter Bonus Feats" & vbCrLf
End If


Case Is = "Monk"
ba = ba + b_att(2, Class(6))
will_save = will_save + will(1, Class(6))
ref_save = ref_save + ref(1, Class(6))
fort_save = fort_save + fort(1, Class(6))
hp = hp + (8 + con_mod) * Class(6)

clslvls = clslvls & Left(select_cls, 4) & "/" & Class(6)

If first_class = select_cls Then
skill_pts = skill_pts + (4 + int_mod) * (Class(6) + 3)
Else
skill_pts = skill_pts + (4 + int_mod) * Class(6)
End If

Select Case Class(6) 'gets Slow fall x ft
Case Is = 20
txtspecial = txtspecial & "Slow Fall: Any Distance" & vbCrLf
Case Is >= 18
txtspecial = txtspecial & "Slow Fall: 90ft" & vbCrLf
Case Is >= 16
txtspecial = txtspecial & "Slow Fall: 80ft" & vbCrLf
Case Is >= 14
txtspecial = txtspecial & "Slow Fall: 70ft" & vbCrLf
Case Is >= 12
txtspecial = txtspecial & "Slow Fall: 60ft" & vbCrLf
Case Is >= 10
txtspecial = txtspecial & "Slow Fall: 50ft" & vbCrLf
Case Is >= 8
txtspecial = txtspecial & "Slow Fall: 40ft" & vbCrLf
Case Is >= 6
txtspecial = txtspecial & "Slow Fall: 30ft" & vbCrLf
Case Is >= 4
txtspecial = txtspecial & "Slow Fall: 20ft" & vbCrLf
End Select

spec_loop = Class(6)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop 'gets non-incremental specials and ac bonus
Case Is = 20
txtspecial = txtspecial & "Perfect Self" & vbCrLf
ac = ac + 1
Case Is = 19
txtspecial = txtspecial & "Empty Body" & vbCrLf
Case Is = 17
txtspecial = txtspecial & "Timeless Body" & vbCrLf & "Tongue of the Sun and Moon" & vbCrLf
Case Is = 16
txtspecial = txtspecial & "Ki Strike(Adamantine)" & vbCrLf
Case Is = 15
txtspecial = txtspecial & "Quivering Palm" & vbCrLf
ac = ac + 1
Case Is = 13
txtspecial = txtspecial & "Diamond Soul" & vbCrLf
Case Is = 12
txtspecial = txtspecial & "Abundant Step" & vbCrLf
Case Is = 11
txtspecial = txtspecial & "Diamond Body" & vbCrLf & "Greater Flurry" & vbCrLf
Case Is = 10
txtspecial = txtspecial & "Ki Strike(Lawful)" & vbCrLf
ac = ac + 1
Case Is = 9
txtspecial = txtspecial & "Improved Evasion" & vbCrLf
Case Is = 7
txtspecial = txtspecial & "Wholeness of Body" & vbCrLf
Case Is = 6
txtspecial = txtspecial & "Bonus Feat - Monk" & vbCrLf
feats = feats + 1
Case Is = 5
txtspecial = txtspecial & "Purity of Body" & vbCrLf
ac = ac + 1
Case Is = 4
txtspecial = txtspecial & "Ki Strike(Magic)" & vbCrLf
Case Is = 3
txtspecial = txtspecial & "Still Mind" & vbCrLf
Case Is = 2
txtspecial = txtspecial & "Evasion" & vbCrLf & "Bonus Feat - Monk" & vbCrLf
feats = feats + 1
Case Is = 1
txtspecial = txtspecial & "Flurry of Blows" & vbCrLf & "Unarmed Strike" & vbCrLf & "Bonus Feat - Monk" & vbCrLf
feats = feats + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Paladin"
ba = ba + b_att(1, Class(7))
will_save = will_save + will(3, Class(7))
ref_save = ref_save + ref(3, Class(7))
fort_save = fort_save + fort(1, Class(7))
hp = hp + (10 + con_mod) * Class(7)

clslvls = clslvls & Left(select_cls, 4) & "/" & Class(7)

If first_class = select_cls Then
skill_pts = skill_pts + (2 = int_mod) * (Class(7) + 3)
Else
skill_pts = skill_pts + (2 + int_mod) * Class(7)
End If

Select Case Class(7) 'gets smite evil/day
Case Is = 20
txtspecial = txtspecial & "Smite Evil 5/Day" & vbCrLf
Case Is >= 15
txtspecial = txtspecial & "Smite Evil 4/Day" & vbCrLf
Case Is >= 10
txtspecial = txtspecial & "Smite Evil 3/Day" & vbCrLf
Case Is >= 5
txtspecial = txtspecial & "Smite Evil 2/Day" & vbCrLf
Case Is >= 1
txtspecial = txtspecial & "Smite Evil 1/Day" & vbCrLf
End Select

Select Case Class(7) 'gets Remove disease/week
Case Is >= 18
txtspecial = txtspecial & "Remove Disease 5/Week" & vbCrLf
Case Is >= 15
txtspecial = txtspecial & "Remove Disease 4/Week" & vbCrLf
Case Is >= 12
txtspecial = txtspecial & "Remove Disease 3/Week" & vbCrLf
Case Is >= 9
txtspecial = txtspecial & "Remove Disease 2/Week" & vbCrLf
Case Is >= 6
txtspecial = txtspecial & "Remove Disease 1/Week" & vbCrLf
End Select

spec_loop = Class(7)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 5
txtspecial = txtspecial & "Special Mount" & vbCrLf
Case Is = 4
txtspecial = txtspecial & "Turn Undead" & vbCrLf
Case Is = 3
txtspecial = txtspecial & "Aura of Courage" & vbCrLf & "Divine Health" & vbCrLf
Case Is = 2
txtspecial = txtspecial & "Divine grace" & vbCrLf & "Lay On Hands" & vbCrLf
Case Is = 1
txtspecial = txtspecial & "Aura of Good" & vbCrLf & "Detect Evil" & vbCrLf
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0

If Class(7) >= 1 Then
txtspells = txtspells & "Paladin Spells/Day - Divine" & vbCrLf
End If

Select Case Class(7) 'gets spells/day
Case Is = 20
txtspells = txtspells & "1:3  2:3  3:3  4:3" & vbCrLf
Case Is = 19
txtspells = txtspells & "1:3  2:3  3:3  4:2" & vbCrLf
Case Is = 18
txtspells = txtspells & "1:3  2:2  3:2  4:1" & vbCrLf
Case Is = 17
txtspells = txtspells & "1:2  2:2  3:2  4:1" & vbCrLf
Case Is = 16
txtspells = txtspells & "1:2  2:2  3:1  4:1" & vbCrLf
Case Is = 15
txtspells = txtspells & "1:2  2:1  3:1  4:1" & vbCrLf
Case Is = 14
txtspells = txtspells & "1:2  2:1  3:1  4:0" & vbCrLf
Case Is = 13
txtspells = txtspells & "1:1  2:1  3:1" & vbCrLf
Case Is = 12
txtspells = txtspells & "1:1  2:1  3:1" & vbCrLf
Case Is = 11
txtspells = txtspells & "1:1  2:1  3:0" & vbCrLf
Case Is = 10
txtspells = txtspells & "1:1  2:1" & vbCrLf
Case Is = 9
txtspells = txtspells & "1:1  2:0" & vbCrLf
Case Is = 8
txtspells = txtspells & "1:1  2:0" & vbCrLf
Case Is = 7
txtspells = txtspells & "1:1" & vbCrLf
Case Is = 6
txtspells = txtspells & "1:1" & vbCrLf
Case Is = 5
txtspells = txtspells & "1:0" & vbCrLf
Case Is = 4
txtspells = txtspells & "1:0" & vbCrLf
End Select


Case Is = "Ranger"
ba = ba + b_att(1, Class(8))
will_save = will_save + will(3, Class(8))
ref_save = ref_save + ref(1, Class(8))
fort_save = fort_save + fort(1, Class(8))
hp = hp + (8 + con_mod) * Class(8)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(8)

If first_class = select_cls Then
skill_pts = skill_pts + (6 + int_mod) * (Class(8) + 3)
Else
skill_pts = skill_pts + (6 + int_mod) * Class(8)
End If

spec_loop = Class(8)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 20
txtspecial = txtspecial & "5th Favoured Enemy" & vbCrL
Case Is = 17
txtspecial = txtspecial & "Hide in Plain Sight" & vbCrLf
Case Is = 15
txtspecial = txtspecial & "4th Favoured Enemy" & vbCrLf
Case Is = 13
txtspecial = txtspecial & " Camouflage" & vbCrLf
Case Is = 11
txtspecial = txtspecial & "Combat Style mastery" & vbCrLf
Case Is = 10
txtspecial = txtspecial & "3rd Favoured Enemy" & vbCrLf
Case Is = 9
txtspecial = txtspecial & "Evasion" & vbCrLf
Case Is = 8
txtspecial = txtspecial & "Swift Tracker" & vbCrLf
Case Is = 7
txtspecial = txtspecial & "Woodland Stride" & vbCrLf
Case Is = 6
txtspecial = txtspecial & "Improved Combat Style" & vbCrLf
Case Is = 5
txtspecial = txtspecial & "2nd Favoured Enemy" & vbCrLf
Case Is = 4
txtspecial = txtspecial & "Animal Companion" & vbCrLf
Case Is = 3
txtspecial = txtspecial & "Endurance" & vbCrLf
Case Is = 2
txtspecial = txtspecial & "Combat Style" & vbCrLf
Case Is = 1
txtspecial = txtspecial & "1st Favoured Enemy" & vbCrLf & "Track" & vbCrLf & "Wild Empathy" & vbCrLf
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0

If Class(8) >= 1 Then
txtspells = txtspells & "Ranger Spells/Day" & vbCrLf
End If

Select Case Class(8) 'gets spells/day
Case Is = 20
txtspells = txtspells & "1:3  2:3  3:3  4:3" & vbCrLf
Case Is = 19
txtspells = txtspells & "1:3  2:3  3:3  4:2" & vbCrLf
Case Is = 18
txtspells = txtspells & "1:3  2:2  3:2  4:1" & vbCrLf
Case Is = 17
txtspells = txtspells & "1:2  2:2  3:2  4:1" & vbCrLf
Case Is = 16
txtspells = txtspells & "1:2  2:2  3:1  4:1" & vbCrLf
Case Is = 15
txtspells = txtspells & "1:2  2:1  3:1  4:1" & vbCrLf
Case Is = 14
txtspells = txtspells & "1:2  2:1  3:1  4:0" & vbCrLf
Case Is = 13
txtspells = txtspells & "1:1  2:1  3:1" & vbCrLf
Case Is = 12
txtspells = txtspells & "1:1  2:1  3:1" & vbCrLf
Case Is = 11
txtspells = txtspells & "1:1  2:1  3:0" & vbCrLf
Case Is = 10
txtspells = txtspells & "1:1  2:1" & vbCrLf
Case Is = 9
txtspells = txtspells & "1:1  2:0" & vbCrLf
Case Is = 8
txtspells = txtspells & "1:1  2:0" & vbCrLf
Case Is = 7
txtspells = txtspells & "1:1" & vbCrLf
Case Is = 6
txtspells = txtspells & "1:1" & vbCrLf
Case Is = 5
txtspells = txtspells & "1:0" & vbCrLf
Case Is = 4
txtspells = txtspells & "1:0" & vbCrLf
Case Is < 4
txtspells = txtspells & "Bonus spells only" & vbCrLf
End Select


Case Is = "Rogue"
ba = ba + b_att(2, Class(9))
will_save = will_save + will(3, Class(9))
ref_save = ref_save + ref(1, Class(9))
fort_save = fort_save + fort(3, Class(9))
hp = hp + (6 + con_mod) * Class(9)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(9)

If first_class = select_cls Then
skill_pts = skill_pts + (8 + int_mod) * (Class(9) + 3)
Else
skill_pts = skill_pts + (8 + int_mod) * Class(9)
End If

Select Case Class(9) 'gets sneak attack +xd6
Case Is >= 19
txtspecial = txtspecial & "Sneak Attack +10d6" & vbCrLf
Case Is >= 17
txtspecial = txtspecial & "Sneak Attack +9d6" & vbCrLf
Case Is >= 15
txtspecial = txtspecial & "Sneak Attack +8d6" & vbCrLf
Case Is >= 13
txtspecial = txtspecial & "Sneak Attack +7d6" & vbCrLf
Case Is >= 11
txtspecial = txtspecial & "Sneak Attack +6d6" & vbCrLf
Case Is >= 9
txtspecial = txtspecial & "Sneak Attack +5d6" & vbCrLf
Case Is >= 7
txtspecial = txtspecial & "Sneak Attack +4d6" & vbCrLf
Case Is >= 5
txtspecial = txtspecial & "Sneak Attack +3d6" & vbCrLf
Case Is >= 3
txtspecial = txtspecial & "Sneak Attack +2d6" & vbCrLf
Case Is >= 1
txtspecial = txtspecial & "Sneak Attack +1d6" & vbCrLf
End Select

Select Case Class(9) 'gets trap sense
Case Is >= 18
txtspecial = txtspecial & "Trap Sense +6" & vbCrLf
Case Is >= 15
txtspecial = txtspecial & "Trap Sense +5" & vbCrLf
Case Is >= 12
txtspecial = txtspecial & "Trap Sense +4" & vbCrLf
Case Is >= 9
txtspecial = txtspecial & "Trap Sense +3" & vbCrLf
Case Is >= 6
txtspecial = txtspecial & "Trap Sense +2" & vbCrLf
Case Is >= 3
txtspecial = txtspecial & "Trap Sense +1" & vbCrLf
End Select

spec_loop = Class(9)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 19
txtspecial = txtspecial & "Special Ability" & vbCrLf
Case Is = 16
txtspecial = txtspecial & "Special Ability" & vbCrLf
Case Is = 13
txtspecial = txtspecial & "Special Ability" & vbCrLf
Case Is = 10
txtspecial = txtspecial & "Special Ability" & vbCrLf
Case Is = 8
txtspecial = txtspecial & "Improved Uncanny Dodge" & vbCrLf
Case Is = 4
txtspecial = txtspecial & "Uncanny Dodge" & vbCrLf
Case Is = 2
txtspecial = txtspecial & "Evasion" & vbCrLf
Case Is = 1
txtspecial = txtspecial & "Trapfinding" & vbCrLf
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Sorcerer"
ba = ba + b_att(3, Class(10))
will_save = will_save + will(1, Class(10))
ref_save = ref_save + ref(3, Class(10))
fort_save = fort_save + fort(3, Class(10))
hp = hp + (4 + con_mod) * Class(10)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(10)

If first_class = select_cls Then
skill_pts = skill_pts + (2 + int_mod) * (Class(10) + 3)
Else
skill_pts = skill_pts + (2 + int_mod) * Class(10)
End If

If Class(10) >= 1 Then
txtspecial = txtspecial & "Summon Familiar" & vbCrLf
End If

If Class(10) >= 1 Then
txtspells = txtspells & "Sorcerer Spells/Day - Arcane" & vbCrLf
End If

Select Case Class(10) 'gets spells/day
Case Is = 20
txtspells = txtspells & "0:6  1:6  2:6  3:6  4:6  5:6  6:6  7:6  8:6  9:6" & vbCrLf
Case Is = 19
txtspells = txtspells & "0:6  1:6  2:6  3:6  4:6  5:6  6:6  7:6  8:6  9:4" & vbCrLf
Case Is = 18
txtspells = txtspells & "0:6  1:6  2:6  3:6  4:6  5:6  6:6  7:6  8:5  9:3" & vbCrLf
Case Is = 17
txtspells = txtspells & "0:6  1:6  2:6  3:6  4:6  5:6  6:6  7:6  8:4" & vbCrLf
Case Is = 16
txtspells = txtspells & "0:6  1:6  2:6  3:6  4:6  5:6  6:6  7:5  8:3 " & vbCrLf
Case Is = 15
txtspells = txtspells & "0:6  1:6  2:6  3:6  4:6  5:6  6:6  7:4" & vbCrLf
Case Is = 14
txtspells = txtspells & "0:6  1:6  2:6  3:6  4:6  5:6  6:5  7:3" & vbCrLf
Case Is = 13
txtspells = txtspells & "0:6  1:6  2:6  3:6  4:6  5:6  6:4" & vbCrLf
Case Is = 12
txtspells = txtspells & "0:6  1:6  2:6  3:6  4:6  5:5  6:3" & vbCrLf
Case Is = 11
txtspells = txtspells & "0:6  1:6  2:6  3:6  4:6  5:4 " & vbCrLf
Case Is = 10
txtspells = txtspells & "0:6  1:6  2:6  3:6  4:5  5:3 " & vbCrLf
Case Is = 9
txtspells = txtspells & "0:6  1:6  2:6  3:6  4:4" & vbCrLf
Case Is = 8
txtspells = txtspells & "0:6  1:6  2:6  3:5  4:3 " & vbCrLf
Case Is = 7
txtspells = txtspells & "0:6  1:6  2:6  3:4" & vbCrLf
Case Is = 6
txtspells = txtspells & "0:6  1:6  2:5  3:3 " & vbCrLf
Case Is = 5
txtspells = txtspells & "0:6  1:6  2:4" & vbCrLf
Case Is = 4
txtspells = txtspells & "0:6  1:6  2:3" & vbCrLf
Case Is = 3
txtspells = txtspells & "0:6  1:5" & vbCrLf
Case Is = 2
txtspells = txtspells & "0:6  1:4" & vbCrLf
Case Is = 1
txtspells = txtspells & "0:5  1:3" & vbCrLf
End Select


Case Is = "Wizard"
ba = ba + b_att(3, Class(11))
will_save = will_save + will(1, Class(11))
ref_save = ref_save + ref(3, Class(11))
fort_save = fort_save + fort(3, Class(11))
hp = hp + (4 + con_mod) * Class(11)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(11)

If first_class = select_cls Then
skill_pts = skill_pts + (2 + int_mod) * (Class(11) + 3)
Else
skill_pts = skill_pts + (2 + int_mod) * Class(11)
End If

spec_loop = Class(11)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 20
txtspecial = txtspecial & "Bonus Feat - Wizard" & vbCrLf
feats = feats + 1
Case Is = 15
txtspecial = txtspecial & "Bonus Feat - Wizard" & vbCrLf
feats = feats + 1
Case Is = 10
txtspecial = txtspecial & "Bonus Feat - Wizard" & vbCrLf
feats = feats + 1
Case Is = 5
txtspecial = txtspecial & "Bonus Feat - Wizard" & vbCrLf
feats = feats + 1
Case Is = 1
txtspecial = txtspecial & "Scribe Scroll" & vbCrLf & "Summon Familiar" & vbCrLf
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0

If Class(11) >= 1 Then
txtspells = txtspells & "Wizard Spells/Day - Arcane" & vbCrLf
End If

Select Case Class(11) 'gets spells/day
Case Is = 20
txtspells = txtspells & "0:4  1:4  2:4  3:4  4:4  5:4  6:4  7:4  8:4  9:4" & vbCrLf
Case Is = 19
txtspells = txtspells & "0:4  1:4  2:4  3:4  4:4  5:4  6:4  7:3  8:3  9:3" & vbCrLf
Case Is = 18
txtspells = txtspells & "0:4  1:4  2:4  3:4  4:4  5:4  6:4  7:2  8:3  9:2" & vbCrLf
Case Is = 17
txtspells = txtspells & "0:4  1:4  2:4  3:4  4:4  5:4  6:4  7:3  8:2  9:1" & vbCrLf
Case Is = 16
txtspells = txtspells & "0:4  1:4  2:4  3:4  4:4  5:4  6:3  7:2  8:2 " & vbCrLf
Case Is = 15
txtspells = txtspells & "0:4  1:4  2:4  3:4  4:4  5:4  6:3  7:3  8:1 " & vbCrLf
Case Is = 14
txtspells = txtspells & "0:4  1:4  2:4  3:4  4:4  5:3  6:3  7:2" & vbCrLf
Case Is = 13
txtspells = txtspells & "0:4  1:4  2:4  3:4  4:4  5:3  6:2  7:1" & vbCrLf
Case Is = 12
txtspells = txtspells & "0:4  1:4  2:4  3:4  4:3  5:3  6:2" & vbCrLf
Case Is = 11
txtspells = txtspells & "0:4  1:4  2:4  3:4  4:3  5:2  6:1" & vbCrLf
Case Is = 10
txtspells = txtspells & "0:4  1:4  2:4  3:3  4:3  5:2 " & vbCrLf
Case Is = 9
txtspells = txtspells & "0:4  1:4  2:4  3:3  4:2  5:1 " & vbCrLf
Case Is = 8
txtspells = txtspells & "0:4  1:4  2:3  3:3  4:2 " & vbCrLf
Case Is = 7
txtspells = txtspells & "0:4  1:4  2:3  3:2  4:1" & vbCrLf
Case Is = 6
txtspells = txtspells & "0:4  1:3  2:3  3:2 " & vbCrLf
Case Is = 5
txtspells = txtspells & "0:4  1:3  2:2  3:1 " & vbCrLf
Case Is = 4
txtspells = txtspells & "0:4  1:3  2:2" & vbCrLf
Case Is = 3
txtspells = txtspells & "0:4  1:2  2:1" & vbCrLf
Case Is = 2
txtspells = txtspells & "0:4  1:2" & vbCrLf
Case Is = 1
txtspells = txtspells & "0:3  1:1" & vbCrLf
End Select


Case Is = "Arcane Archer"
ba = ba + b_att(1, Class(12))
will_save = will_save + will(2, Class(12))
ref_save = ref_save + ref(1, Class(12))
fort_save = fort_save + fort(1, Class(12))
hp = hp + (8 + con_mod) * Class(12)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(12)
skill_pts = skill_pts + (4 + int_mod) * Class(12)

Select Case Class(12) 'gets enhance arrow +x
Case Is >= 9
txtspecial = txtspecial & "Enhance Arrow +5" & vbCrLf
Case Is >= 7
txtspecial = txtspecial & "Enhance Arrow +4" & vbCrLf
Case Is >= 5
txtspecial = txtspecial & "Enhance Arrow +3" & vbCrLf
Case Is >= 3
txtspecial = txtspecial & "Enhance Arrow +2" & vbCrLf
Case Is >= 1
txtspecial = txtspecial & "Enhance Arrow +1" & vbCrLf
End Select

spec_loop = Class(12)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Arrow of Death" & vbCrLf
Case Is = 8
txtspecial = txtspecial & "Hail of Arrows" & vbCrLf
Case Is = 6
txtspecial = txtspecial & "Phase Arrow" & vbCrLf
Case Is = 4
txtspecial = txtspecial & "Seeker Arrow" & vbCrLf
Case Is = 2
txtspecial = txtspecial & "Imbue Arrow" & vbCrLf
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Arcane Trickster"
ba = ba + b_att(3, Class(13))
will_save = will_save + will(1, Class(13))
ref_save = ref_save + ref(1, Class(13))
fort_save = fort_save + fort(3, Class(13))
hp = hp + 4 * Class(13)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(13)
skill_pts = skill_pts + (4 + int_mod) * Class(13)

Select Case Class(13) 'gets sneak attack +xd6
Case Is = 10
txtspecial = txtspecial & "Sneak Attack +5d6" & vbCrLf
Case Is >= 8
txtspecial = txtspecial & "Sneak Attack +4d6" & vbCrLf
Case Is >= 6
txtspecial = txtspecial & "Sneak Attack +3d6" & vbCrLf
Case Is >= 4
txtspecial = txtspecial & "Sneak Attack +2d6" & vbCrLf
Case Is >= 2
txtspecial = txtspecial & "Sneak Attack +1d6" & vbCrLf
End Select

Select Case Class(13) 'gets ranged legerdemain/day
Case Is >= 9
txtspecial = txtspecial & "Ranged Legerdemain 3/Day" & vbCrLf
Case Is >= 5
txtspecial = txtspecial & "Ranged Legerdemain 2/Day" & vbCrLf
Case Is >= 1
txtspecial = txtspecial & "Ranged Legerdemain 1/Day" & vbCrLf
End Select

Select Case Class(13) 'gets impromptu sneak attack/day
Case Is >= 7
txtspecial = txtspecial & "Impromptu Sneak Attack 2/Day" & vbCrLf
Case Is >= 3
txtspecial = txtspecial & "Impromptu Sneak Attack 1/Day" & vbCrLf
End Select

txtspells = txtspells & "spells/day gained as if " & Class(13) & " levels were taken in a previously taken casting class."


Case Is = "Archmage"
ba = ba + b_att(3, Class(14))
will_save = will_save + will(1, Class(14))
ref_save = ref_save + ref(3, Class(14))
fort_save = fort_save + fort(3, Class(14))
hp = hp + (4 + con_mod) * Class(14)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(14)
skill_pts = skill_pts + (2 + int_mod) * Class(14)

spec_loop = Class(14)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 5
txtspecial = txtspecial & "High Arcana - " & vbCrLf
Case Is = 4
txtspecial = txtspecial & "High Arcana - " & vbCrLf
Case Is = 3
txtspecial = txtspecial & "High Arcana - " & vbCrLf
Case Is = 2
txtspecial = txtspecial & "High Arcana - " & vbCrLf
Case Is = 1
txtspecial = txtspecial & "High Arcana - " & vbCrLf
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0

txtspells = txtspells & "spells/day gained as if " & Class(14) & " levels were taken in existing arcane casting class."


Case Is = "Assassin"
ba = ba + b_att(2, Class(15))
will_save = will_save + will(3, Class(15))
ref_save = ref_save + ref(1, Class(15))
fort_save = fort_save + fort(3, Class(15))
hp = hp + (6 + con_mod) * Class(15)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(15)
skill_pts = skill_pts + (4 + int_mod) * Class(15)

Select Case Class(15) 'gets sneak attack +xd6
Case Is >= 9
txtspecial = txtspecial & "Sneak Attack +5d6" & vbCrLf
Case Is >= 7
txtspecial = txtspecial & "Sneak Attack +4d6" & vbCrLf
Case Is >= 5
txtspecial = txtspecial & "Sneak Attack +3d6" & vbCrLf
Case Is >= 3
txtspecial = txtspecial & "Sneak Attack +2d6" & vbCrLf
Case Is >= 1
txtspecial = txtspecial & "Sneak Attack +1d6" & vbCrLf
End Select

Select Case Class(15) 'gets +x save against poison
Case Is = 10
txtspecial = txtspecial & "+5 Save Against Poison" & vbCrLf
Case Is >= 8
txtspecial = txtspecial & "+5 Save Against Poison" & vbCrLf
Case Is >= 6
txtspecial = txtspecial & "+5 Save Against Poison" & vbCrLf
Case Is >= 4
txtspecial = txtspecial & "+5 Save Against Poison" & vbCrLf
Case Is >= 2
txtspecial = txtspecial & "+5 Save Against Poison" & vbCrLf
End Select

spec_loop = Class(15)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 8
txtspecial = txtspecial & "Hide in Plain Sight" & vbCrLf
Case Is = 5
txtspecial = txtspecial & "Improved Uncanny Dodge" & vbCrLf
Case Is = 2
txtspecial = txtspecial & "Uncanny Dodge" & vbCrLf
Case Is = 1
txtspecial = txtspecial & "Death Attack" & vbCrLf & "Poison Use" & vbCrLf & "spells" & vbCrLf
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0

If Class(15) >= 1 Then
txtspells = txtspells & "Assassin Spells/Day - Arcane" & vbCrLf
End If

Select Case Class(15) 'gets arcane spells
Case Is = 10
txtspells = txtspells & "1:3  2:3  3:3  4:3" & vbCrLf
Case Is = 9
txtspells = txtspells & "1:3  2:3  3:3  4:2" & vbCrLf
Case Is = 8
txtspells = txtspells & "1:3  2:3  3:3  4:1" & vbCrLf
Case Is = 7
txtspells = txtspells & "1:3  2:3  3:2  4:0" & vbCrLf
Case Is = 6
txtspells = txtspells & "1:3  2:3  3:1" & vbCrLf
Case Is = 5
txtspells = txtspells & "1:3  2:2  3:0" & vbCrLf
Case Is = 4
txtspells = txtspells & "1:3  2:1" & vbCrLf
Case Is = 3
txtspells = txtspells & "1:2  2:0" & vbCrLf
Case Is = 2
txtspells = txtspells & "1:1" & vbCrLf
Case Is = 1
txtspells = txtspells & "1:0" & vbCrLf
End Select


Case Is = "Blackguard"
ba = ba + b_att(1, Class(16))
will_save = will_save + will(3, Class(16))
ref_save = ref_save + ref(3, Class(16))
fort_save = fort_save + fort(1, Class(16))
hp = hp + (10 + con_mod) * Class(16)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(16)
skill_pts = skill_pts + (2 + int_mod) * Class(16)

Select Case Class(16) 'gets sneak attack +xd6
Case Is = 10
txtspecial = txtspecial & "Sneak Attack +3d6" & vbCrLf
Case Is >= 7
txtspecial = txtspecial & "Sneak Attack +2d6" & vbCrLf
Case Is >= 4
txtspecial = txtspecial & "Sneak Attack +1d6" & vbCrLf
End Select

Select Case Class(16) 'gets smite good/day
Case Is = 10
txtspecial = txtspecial & "Smite Good 3/Day" & vbCrLf
Case Is >= 5
txtspecial = txtspecial & "Smite Good 2/Day" & vbCrLf
Case Is >= 2
txtspecial = txtspecial & "Smite Good 1/Day" & vbCrLf
End Select

spec_loop = Class(16)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 5
txtspecial = txtspecial & "Fiendish Servant" & vbCrLf
Case Is = 3
txtspecial = txtspecial & "Command Undead" & vbCrLf & "Aura of Despair" & vbCrLf
Case Is = 2
txtspecial = txtspecial & "Dark Blessing" & vbCrLf
Case Is = 1
txtspecial = txtspecial & "Aura of Evil" & vbCrLf & "Detect Good" & vbCrLf & "Poison Use" & vbCrLf
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0

If Class(16) >= 1 Then
txtspells = txtspells & "Blackguard Spells/Day - Divine" & vbCrLf
End If

Select Case Class(16) 'gets arcane spells
Case Is = 10
txtspells = txtspells & "1:2  2:2  3:2  4:1" & vbCrLf
Case Is = 9
txtspells = txtspells & "1:2  2:2  3:1  4:1" & vbCrLf
Case Is = 8
txtspells = txtspells & "1:2  2:1  3:1  4:1" & vbCrLf
Case Is = 7
txtspells = txtspells & "1:2  2:1  3:1  4:0" & vbCrLf
Case Is = 6
txtspells = txtspells & "1:1  2:1  3:1" & vbCrLf
Case Is = 5
txtspells = txtspells & "1:1  2:1  3:0" & vbCrLf
Case Is = 4
txtspells = txtspells & "1:1  2:1" & vbCrLf
Case Is = 3
txtspells = txtspells & "1:1  2:0" & vbCrLf
Case Is = 2
txtspells = txtspells & "1:1" & vbCrLf
Case Is = 1
txtspells = txtspells & "1:0" & vbCrLf
End Select


Case Is = "Dragon Disciple"
ba = ba + b_att(2, Class(17))
will_save = will_save + will(1, Class(17))
ref_save = ref_save + ref(3, Class(17))
fort_save = fort_save + fort(1, Class(17))
hp = hp + (12 + con_mod) * Class(17)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(17)
skill_pts = skill_pts + (2 + int_mod) * Class(17)

Select Case Class(17)
Case Is = 10
txtspecial = txtspecial & "Breath Weapon 2d8" & vbCrLf
Case Is >= 7
txtspecial = txtspecial & "Breath Weapon 4d8" & vbCrLf
Case Is >= 3
txtspecial = txtspecial & "Breath Weapon 6d8" & vbCrLf
End Select

Select Case Class(17)
Case Is = 10
txtspecial = txtspecial & "Blindsense 60ft" & vbCrLf
Case Is >= 5
txtspecial = txtspecial & "Blindsense 30ft" & vbCrLf
End Select

spec_loop = Class(17)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Dragon Apotheosis" & vbCrLf
ac = ac + 1
strength = strength + 4
charisma = charisma + 2
txtspecial = txtspecial & "lowlight vision" & vbCrLf
txtspecial = txtspecial & "immune to sleep and paralysis" & vbCrLf
txtspecial = txtspecial & "immune to breath weapon type" & vbCrLf
Case Is = 9
txtspecial = txtspecial & "Wings" & vbCrLf
Case Is = 8
intelligence = intelligence + 2
skill_pts = skill_pts + level 'additional skill points as int increases. this doesn't affect anything else, as the new modifiers are calculated after all classes are taken
Case Is = 7
ac = ac + 1
Case Is = 6
constitution = constitution + 2
hp = hp + level 'additional hp as con increases. this doesn't affect anything else, as the new modifiers are calculated after all classes are taken
Case Is = 4
strength = strength + 2
ac = ac + 1
Case Is = 3
Case Is = 2
txtspecial = txtspecial & "Claws & Bite" & vbCrLf
strength = strength + 2
Case Is = 1
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0

Select Case Class(17) 'bonus spells
Case Is >= 9
txtspecial = txtspecial & "7 bonus spells" & vbCrLf
Case Is >= 8
txtspecial = txtspecial & "6 bonus spells" & vbCrLf
Case Is >= 6
txtspecial = txtspecial & "5 bonus spells" & vbCrLf
Case Is >= 5
txtspecial = txtspecial & "4 bonus spells" & vbCrLf
Case Is >= 4
txtspecial = txtspecial & "3 bonus spells" & vbCrLf
Case Is >= 2
txtspecial = txtspecial & "2 bonus spells" & vbCrLf
Case Is >= 1
txtspecial = txtspecial & "1 bonus spells" & vbCrLf
End Select


Case Is = "Duelist"
ba = ba + b_att(1, Class(18))
will_save = will_save + will(3, Class(18))
ref_save = ref_save + ref(1, Class(18))
fort_save = fort_save + fort(3, Class(18))
hp = hp + (10 + con_mod) * Class(18)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(18)
skill_pts = skill_pts + (4 + int_mod) * Class(18)

Select Case Class(18) 'gets improved reaction
Case Is >= 8
txtspecial = txtspecial & "" & vbCrLf
Case Is >= 2
txtspecial = txtspecial & "" & vbCrLf
End Select

Select Case Class(18) 'gets precise strike +xd6
Case Is = 10
txtspecial = txtspecial & "Precise Strike +2d6" & vbCrLf
Case Is >= 5
txtspecial = txtspecial & "Precise Strike +1d6" & vbCrLf
End Select

spec_loop = Class(18)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 9
txtspecial = txtspecial & "Deflect Arrows" & vbCrLf
Case Is = 7
txtspecial = txtspecial & "Elaborate Parry" & vbCrLf
Case Is = 6
txtspecial = txtspecial & "Acrobatic Charge" & vbCrLf
Case Is = 4
txtspecial = txtspecial & "Grace" & vbCrLf
Case Is = 3
txtspecial = txtspecial & "Enhanced Mobility" & vbCrLf
Case Is = 1
txtspecial = txtspecial & "Canny Defense" & vbCrLf
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Dwarven Defender"
ba = ba + b_att(1, Class(19))
will_save = will_save + will(1, Class(19))
ref_save = ref_save + ref(3, Class(19))
fort_save = fort_save + fort(1, Class(19))
hp = hp + (12 + con_mod) * Class(19)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(19)
skill_pts = skill_pts + (2 + int_mod) * Class(19)

Select Case Class(19) 'gets defensive stance/day
Case Is >= 9
txtspecial = txtspecial & "Defensive Stance 5/Day" & vbCrLf
Case Is >= 7
txtspecial = txtspecial & "Defensive Stance 4/Day" & vbCrLf
Case Is >= 5
txtspecial = txtspecial & "Defensive Stance 3/Day" & vbCrLf
Case Is >= 3
txtspecial = txtspecial & "Defensive Stance 2/Day" & vbCrLf
Case Is >= 1
txtspecial = txtspecial & "Defensive Stance 1/Day" & vbCrLf
End Select

Select Case Class(19) 'gets damage reduction
Case Is = 10
txtspecial = txtspecial & "Damage Reduction 6/-" & vbCrLf
Case Is >= 6
txtspecial = txtspecial & "Damage Reduction 3/-" & vbCrLf
End Select

spec_loop = Class(19)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
ac = ac + 1
Case Is = 8
txtspecial = txtspecial & "Mobile Defense" & vbCrLf
Case Is = 7
ac = ac + 1
Case Is = 6
txtspecial = txtspecial & "Improved Uncanny Dodge" & vbCrLf
Case Is = 4
ac = ac + 1
Case Is = 2
txtspecial = txtspecial & "Uncanny Dodge" & vbCrLf
Case Is = 1
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Eldritch Knight"
ba = ba + b_att(1, Class(20))
will_save = will_save + will(3, Class(20))
ref_save = ref_save + ref(3, Class(20))
fort_save = fort_save + fort(1, Class(20))
hp = hp + (6 + con_mod) * Class(20)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(20)
skill_pts = skill_pts + (6 + int_mod) * Class(20)

If Class(20) >= 1 Then
txtspecial = txtspecial & "Fighter Bonus Feat" & vbCrLf
feats = feats + 1
End If

txtspells = txtspells & "spells/day gained as if " & Class(20) & " levels were taken in existing arcane casting class."


Case Is = "Hierophant"
ba = ba + b_att(3, Class(21))
will_save = will_save + will(1, Class(21))
ref_save = ref_save + ref(3, Class(21))
fort_save = fort_save + fort(1, Class(21))
hp = hp + (8 + con_mod) * Class(21)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(21)
skill_pts = skill_pts + (2 + int_mod) * Class(21)

spec_loop = Class(21)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 5
txtspecial = txtspecial & "Special Ability - " & vbCrLf
Case Is = 4
txtspecial = txtspecial & "Special Ability - " & vbCrLf
Case Is = 3
txtspecial = txtspecial & "Special Ability - " & vbCrLf
Case Is = 2
txtspecial = txtspecial & "Special Ability - " & vbCrLf
Case Is = 1
txtspecial = txtspecial & "Special Ability - " & vbCrLf
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Horizon Walker"
ba = ba + b_att(1, Class(22))
will_save = will_save + will(3, Class(22))
ref_save = ref_save + ref(3, Class(22))
fort_save = fort_save + fort(1, Class(22))
hp = hp + (8 + con_mod) * Class(22)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(22)
skill_pts = skill_pts + (6 + int_mod) * Class(22)

spec_loop = Class(22)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Planar Terrain Mastery - " & vbCrLf
Case Is = 9
txtspecial = txtspecial & "Planar Terrain Mastery - " & vbCrLf
Case Is = 8
txtspecial = txtspecial & "Planar Terrain Mastery - " & vbCrLf
Case Is = 7
txtspecial = txtspecial & "Planar Terrain Mastery - " & vbCrLf
Case Is = 6
txtspecial = txtspecial & "Planar Terrain Mastery - " & vbCrLf
Case Is = 5
txtspecial = txtspecial & "Terrain Mastery - " & vbCrLf
Case Is = 4
txtspecial = txtspecial & "Terrain Mastery - " & vbCrLf
Case Is = 3
txtspecial = txtspecial & "Terrain Mastery - " & vbCrLf
Case Is = 2
txtspecial = txtspecial & "Terrain Mastery - " & vbCrLf
Case Is = 1
txtspecial = txtspecial & "Terrain Mastery - " & vbCrLf
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Loremaster"
ba = ba + b_att(3, Class(3))
will_save = will_save + will(1, Class(23))
ref_save = ref_save + ref(3, Class(23))
fort_save = fort_save + fort(1, Class(23))
hp = hp + (4 + con_mod) * Class(23)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(23)
skill_pts = skill_pts + (4 + int_mod) * Class(23)

spec_loop = Class(23)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "True Lore" & vbCrLf
Case Is = 9
txtspecial = txtspecial & "Secret - " & vbCrLf
Case Is = 8
txtspecial = txtspecial & "Bonus Language" & vbCrLf
Case Is = 7
txtspecial = txtspecial & "Secret - " & vbCrLf
Case Is = 6
txtspecial = txtspecial & "Greater Lore" & vbCrLf
Case Is = 5
txtspecial = txtspecial & "Secret - " & vbCrLf
Case Is = 4
txtspecial = txtspecial & "Bonus Language" & vbCrLf
Case Is = 3
txtspecial = txtspecial & "Secret - " & vbCrLf
Case Is = 2
txtspecial = txtspecial & "Lore" & vbCrLf
Case Is = 1
txtspecial = txtspecial & "Secret - " & vbCrLf
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0

txtspells = txtspells & "spells/day gained as if " & Class(23) & " levels were taken in a previously taken casting class."


Case Is = "Mystic Theurge"
ba = ba + b_att(3, Class(24))
will_save = will_save + will(1, Class(24))
ref_save = ref_save + ref(3, Class(24))
fort_save = fort_save + fort(1, Class(24))
hp = hp + (4 + con_mod) * Class(24)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(24)
skill_pts = skill_pts + (2 + int_mod) * Class(24)

txtspells = txtspells & "spells/day gained as if " & Class(24) & " levels were taken in existing arcane casting class."
txtspells = txtspells & "spells/day gained as if " & Class(24) & " levels were taken in existing divine casting class."


Case Is = "Shadowdancer"
ba = ba + b_att(2, Class(25))
will_save = will_save + will(3, Class(25))
ref_save = ref_save + ref(1, Class(25))
fort_save = fort_save + fort(3, Class(25))
hp = hp + (8 + con_mod) * Class(25)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(25)
skill_pts = skill_pts + (6 + int_mod) * Class(25)

Select Case Class(25) 'gets shadow jump xft
Case Is = 10
txtspecial = txtspecial & "Shadow Jump 160ft" & vbCrLf
Case Is >= 8
txtspecial = txtspecial & "Shadow Jump 80ft" & vbCrLf
Case Is >= 6
txtspecial = txtspecial & "Shadow Jump 40ft" & vbCrLf
Case Is >= 4
txtspecial = txtspecial & "Shadow Jump 20ft" & vbCrLf
End Select

spec_loop = Class(25)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Improved Evasion" & vbCrLf
Case Is = 9
txtspecial = txtspecial & "Summon Shadow" & vbCrLf
Case Is = 7
txtspecial = txtspecial & "Slippery Mind" & vbCrLf
Case Is = 6
txtspecial = txtspecial & "Summon Shadow" & vbCrLf
Case Is = 5
txtspecial = txtspecial & "Defensive Roll" & vbCrLf & "Improved Uncanny Dodge" & vbCrLf
Case Is = 3
txtspecial = txtspecial & "Shadow Illusion" & vbCrLf & "Summon Shadow" & vbCrLf
Case Is = 2
txtspecial = txtspecial & "Evasion" & vbCrLf & "Darkvision" & vbCrLf & "Uncanny Dodge" & vbCrLf
Case Is = 1
txtspecial = txtspecial & "Hide in Plain Sight" & vbCrLf
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Thaumaturgist"
ba = ba + b_att(3, Class(26))
will_save = will_save + will(1, Class(26))
ref_save = ref_save + ref(3, Class(26))
fort_save = fort_save + fort(3, Class(26))
hp = hp + (4 + con_mod) * Class(26)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(26)
skill_pts = skill_pts + (2 + int_mod) * Class(26)

spec_loop = Class(26)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 5
txtspecial = txtspecial & "Planar Cohort" & vbCrLf
Case Is = 4
txtspecial = txtspecial & "Contingent Conjuration" & vbCrLf
Case Is = 3
txtspecial = txtspecial & "Extended Summoning" & vbCrLf
Case Is = 2
txtspecial = txtspecial & "Augment Summoning" & vbCrLf
Case Is = 1
txtspecial = txtspecial & "Improved Ally" & vbCrLf
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0

txtspells = txtspells & "Spells/day gained as if " & Class(26) & " levels were taken in a previously taken casting class."

End Select

Next increment

counter(3) = counter(3) + 1

Loop Until select_cls = ""

txtmods = "BA" & vbTab & ba & vbCrLf & "HP" & vbTab & hp & vbCrLf & "WILL" & vbTab & will_save & vbCrLf & "FORT" & vbTab & fort_save & vbCrLf & "REF" & vbTab & ref_save
frmcharsheet.txtmisc = frmcharsheet.txtmisc & "Classes" & vbTab & vbTab & clslvls
End Sub

Sub mod_classes(ByVal level As Integer, ByRef ba As Integer, ByRef b_att() As Integer, ByRef will_save As Integer, ByRef will() As Integer, ByRef ref_save As Integer, ByRef ref() As Integer, ByRef fort_save As Integer, ByRef fort() As Integer, ByRef hp As Integer, ByVal con_mod As Integer, ByRef skill_pts As Integer, ByVal int_mod As Integer, ByRef ac As Integer, ByRef feats As Integer)

Dim counter(3) As Integer
Dim Class(18) As Integer 'finds out the levels in each class
Dim lvl As Integer 'individual level in class
Dim ttl_lvl As Integer 'total levels used
Dim cls As String
Dim chosen(18) As String 'classes with levels in them
Dim selected As Integer
Dim select_cls As String
Dim increment As Integer
Dim first_class As String

ttl_lvl = 0
counter(2) = 1
counter(3) = 1

Do
For counter(1) = 1 To level
If ttl_lvl < level Then

If ttl_lvl < 3 Then 'makes sure that first 3 levels are not prestige
Do
cls = InputBox("Select the class that you wish to take. Initial letter must be uppercase, i.e. Strong, Fast. Can not be prestige.")
Loop Until cls = "Strong" Or cls = "Fast" Or cls = "Tough" Or cls = "Smart" Or cls = "Dedicated" Or cls = "Charismatic"

Else 'if more than 3 levels have been taken, prestige is allowed.
Do
cls = InputBox("Select the class that you wish to take. Initial letter must be uppercase, i.e. Strong, Fast, Martial Artist.")
Loop Until cls = "Strong" Or cls = "Fast" Or cls = "Tough" Or cls = "Smart" Or cls = "Dedicated" Or cls = "Charismatic" Or cls = "Soldier" Or cls = "Martial Artist" Or cls = "Gunslinger" Or cls = "Infiltrator" Or cls = "Daredevil" Or cls = "Bodyguard" Or cls = "Field Scientist" Or cls = "Techie" Or cls = "Field Medic" Or cls = "Investigator" Or cls = "Personality" Or cls = "Negotiator"
End If

If counter(1) = 1 Then
first_class = cls
End If

chosen(counter(2)) = cls

Do
lvl = InputBox("Select how many levels you wish to take in that class. You have " & level - ttl_lvl & " levels left to spend.")

'makes sure that you can't take over the maximum amount of levels in a class

If cls = "Strong" Or cls = "Fast" Or cls = "Tough" Or cls = "Smart" Or cls = "Dedicated" Or cls = "Charismatic" Or cls = "Soldier" Or cls = "Martial Artist" Or cls = "Gunslinger" Or cls = "Infiltrator" Or cls = "Daredevil" Or cls = "Bodyguard" Or cls = "Field Scientist" Or cls = "Techie" Or cls = "Field Medic" Or cls = "Investigator" Or cls = "Personality" Or cls = "Negotiator" Then

Do
If lvl > 10 Or lvl < 0 Then
lvl = InputBox("Max of 10 levels in class you wish to take in that class.")
End If
Loop Until lvl <= 10
End If

Loop Until lvl <= level 'the user can't continue if he has more than his maximum level
ttl_lvl = ttl_lvl + lvl



Select Case cls 'finds out how many levels are in each class
Case Is = "Strong"
Class(1) = Class(1) + lvl
If Class(1) > 10 Then
Class(1) = Class(1) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Fast"
Class(2) = Class(2) + lvl
If Class(2) > 10 Then
Class(2) = Class(2) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Tough"
Class(3) = Class(3) + lvl
If Class(3) > 10 Then
Class(3) = Class(3) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Smart"
Class(4) = Class(4) + lvl
If Class(4) > 10 Then
Class(4) = Class(4) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Dedicated"
Class(5) = Class(5) + lvl
If Class(5) > 10 Then
Class(5) = Class(5) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Charismatic"
Class(6) = Class(6) + lvl
If Class(6) > 10 Then
Class(6) = Class(6) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Soldier"
Class(7) = Class(7) + lvl
If Class(7) > 10 Then
Class(7) = Class(7) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Martial Artist"
Class(8) = Class(8) + lvl
If Class(8) > 10 Then
Class(8) = Class(8) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Gunslinger"
Class(9) = Class(9) + lvl
If Class(9) > 10 Then
Class(9) = Class(9) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Infiltrator"
Class(10) = Class(10) + lvl
If Class(10) > 10 Then
Class(10) = Class(10) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Daredevil"
Class(11) = Class(11) + lvl
If Class(11) > 10 Then
Class(11) = Class(11) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Bodyguard"
Class(12) = Class(12) + lvl
If Class(12) > 10 Then
Class(12) = Class(12) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Field Scientist"
Class(13) = Class(13) + lvl
If Class(13) > 10 Then
Class(13) = Class(13) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Techie"
Class(14) = Class(14) + lvl
If Class(14) > 10 Then
Class(14) = Class(14) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Field Medic"
Class(15) = Class(15) + lvl
If Class(15) > 10 Then
Class(15) = Class(15) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Investigator"
Class(16) = Class(16) + lvl
If Class(16) > 10 Then
Class(16) = Class(16) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Personality"
Class(17) = Class(17) + lvl
If Class(17) > 10 Then
Class(17) = Class(17) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Negotiator"
Class(18) = Class(18) + lvl
If Class(18) > 10 Then
Class(18) = Class(8) - lvl
MsgBox "Class can not have that many levels in it."
End If

End Select


counter(2) = counter(2) + 1

End If

Next counter(1)

Loop Until ttl_lvl = level


increment = 1

Do 'makes sure that no duplicates appear in chosen class

select_cls = chosen(increment)

For increment = 1 To counter(2)
selected = increment + 1
If select_cls = chosen(selected) Then
chosen(selected) = "x"
End If
Next increment

Dim spec_loop As Integer 'gets the specials from classes
Dim clslvls As String 'puts down how many levels are in each class, which will be put into the misc area afterwards

increment = 1

For increment = 1 To counter(2)
select_cls = chosen(increment)

Select Case select_cls

Case Is = "Strong"
ba = ba + b_att(1, Class(1))
will_save = will_save + will(3, Class(1))
ref_save = ref_save + ref(3, Class(1))
fort_save = fort_save + fort(2, Class(1))
hp = hp + (8 + con_mod) * Class(1)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(1)

If first_class = select_cls Then
skill_pts = skill_pts + (2 + int_mod) * (Class(1) + 3)
Else
skill_pts = skill_pts + (2 + int_mod) * Class(1)
End If

spec_loop = Class(1)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Bonus Feat - Strong" & vbCrLf
feats = feats + 1
Case Is = 9
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
ac = ac + 1
Case Is = 8
txtspecial = txtspecial & "Bonus Feat - Strong" & vbCrLf
feats = feats + 1
Case Is = 7
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
ac = ac + 1
Case Is = 6
txtspecial = txtspecial & "Bonus Feat - Strong" & vbCrLf
feats = feats + 1
Case Is = 5
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
Case Is = 4
txtspecial = txtspecial & "Bonus Feat - Strong" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 3
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
Case Is = 2
txtspecial = txtspecial & "Bonus Feat - Strong" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 1
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Fast"
ba = ba + b_att(2, Class(2))
will_save = will_save + will(3, Class(2))
ref_save = ref_save + ref(2, Class(2))
fort_save = fort_save + fort(3, Class(2))
hp = hp + (8 + con_mod) * Class(2)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(2)

If first_class = select_cls Then
skill_pts = skill_pts + (4 + int_mod) * (Class(2) + 3)
Else
skill_pts = skill_pts + (4 + int_mod) * Class(2)
End If

spec_loop = Class(2)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Bonus Feat - Fast" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 9
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
Case Is = 8
txtspecial = txtspecial & "Bonus Feat - Fast" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 7
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
Case Is = 6
txtspecial = txtspecial & "Bonus Feat - Fast" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 5
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
Case Is = 4
txtspecial = txtspecial & "Bonus Feat - Fast" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 3
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
Case Is = 2
txtspecial = txtspecial & "Bonus Feat - Fast" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 1
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
ac = ac + 3
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Tough"
ba = ba + b_att(2, Class(3))
will_save = will_save + will(3, Class(3))
ref_save = ref_save + ref(3, Class(3))
fort_save = fort_save + fort(2, Class(3))
hp = hp + (10 + con_mod) * Class(3)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(3)

If first_class = select_cls Then
skill_pts = skill_pts + (2 + int_mod) * (Class(3) + 3)
Else
skill_pts = skill_pts + (2 + int_mod) * Class(3)
End If

spec_loop = Class(3)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Bonus Feat - Tough" & vbCrLf
feats = feats + 1
Case Is = 9
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
ac = ac + 1
Case Is = 8
txtspecial = txtspecial & "Bonus Feat - Tough" & vbCrLf
feats = feats + 1
Case Is = 7
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
ac = ac + 1
Case Is = 6
txtspecial = txtspecial & "Bonus Feat - Tough" & vbCrLf
feats = feats + 1
Case Is = 5
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
Case Is = 4
txtspecial = txtspecial & "Bonus Feat - Tough" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 3
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
Case Is = 2
txtspecial = txtspecial & "Bonus Feat - Tough" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 1
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Smart"
ba = ba + b_att(3, Class(4))
will_save = will_save + will(2, Class(4))
ref_save = ref_save + ref(3, Class(4))
fort_save = fort_save + fort(3, Class(4))
hp = hp + (6 + con_mod) * Class(4)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(4)

If first_class = select_cls Then
skill_pts = skill_pts + (8 = int_mod) * (Class(4) + 3)
Else
skill_pts = skill_pts + (8 + int_mod) * Class(4)
End If

spec_loop = Class(4)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Bonus Feat - Smart" & vbCrLf
feats = feats + 1
Case Is = 9
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
Case Is = 8
txtspecial = txtspecial & "Bonus Feat - Smart" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 7
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
Case Is = 6
txtspecial = txtspecial & "Bonus Feat - Smart" & vbCrLf
feats = feats + 1
Case Is = 5
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
ac = ac + 1
Case Is = 4
txtspecial = txtspecial & "Bonus Feat - Smart" & vbCrLf
feats = feats + 1
Case Is = 3
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
Case Is = 2
txtspecial = txtspecial & "Bonus Feat - Smart" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 1
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Dedicated"
ba = ba + b_att(2, Class(5))
will_save = will_save + will(2, Class(5))
ref_save = ref_save + ref(3, Class(5))
fort_save = fort_save + fort(2, Class(5))
hp = hp + (6 + con_mod) * Class(5)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(5)

If first_class = select_cls Then
skill_pts = skill_pts + (4 = int_mod) * (Class(5) + 3)
Else
skill_pts = skill_pts + (4 + int_mod) * Class(5)
End If

spec_loop = Class(5)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Bonus Feat - Dedicated" & vbCrLf
feats = feats + 1
Case Is = 9
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
ac = ac + 1
Case Is = 8
txtspecial = txtspecial & "Bonus Feat - Dedicated" & vbCrLf
feats = feats + 1
Case Is = 7
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
ac = ac + 1
Case Is = 6
txtspecial = txtspecial & "Bonus Feat - Dedicated" & vbCrLf
feats = feats + 1
Case Is = 5
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
Case Is = 4
txtspecial = txtspecial & "Bonus Feat - Dedicated" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 3
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
Case Is = 2
txtspecial = txtspecial & "Bonus Feat - Dedicated" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 1
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Charismatic"
ba = ba + b_att(3, Class(6))
will_save = will_save + will(3, Class(6))
ref_save = ref_save + ref(2, Class(6))
fort_save = fort_save + fort(2, Class(6))
hp = hp + (6 + con_mod) * Class(6)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(6)

If first_class = select_cls Then
skill_pts = skill_pts + (6 + int_mod) * (Class(6) + 3)
Else
skill_pts = skill_pts + (6 + int_mod) * Class(6)
End If

spec_loop = Class(6)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Bonus Feat - Charismatic" & vbCrLf
feats = feats + 1
Case Is = 9
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
Case Is = 8
txtspecial = txtspecial & "Bonus Feat - Charismatic" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 7
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
Case Is = 6
txtspecial = txtspecial & "Bonus Feat - Charismatic" & vbCrLf
feats = feats + 1
Case Is = 5
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
ac = ac + 1
Case Is = 4
txtspecial = txtspecial & "Bonus Feat - Charismatic" & vbCrLf
feats = feats + 1
Case Is = 3
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
Case Is = 2
txtspecial = txtspecial & "Bonus Feat - Charismatic" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 1
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Soldier"
ba = ba + b_att(1, Class(7))
will_save = will_save + will(3, Class(7))
ref_save = ref_save + ref(2, Class(7))
fort_save = fort_save + fort(2, Class(7))
hp = hp + (10 + con_mod) * Class(7)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(7)
skill_pts = skill_pts + (4 + int_mod) * Class(7)

spec_loop = Class(7)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Critical Strike" & vbCrLf
Case Is = 9
txtspecial = txtspecial & "Bonus Feat - Soldier" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 8
txtspecial = txtspecial & "Greater Weapon Specialisation" & vbCrLf
Case Is = 7
txtspecial = txtspecial & "Improved Reaction" & vbCrLf
ac = ac + 1
Case Is = 6
txtspecial = txtspecial & "Bonus Feat - Soldier" & vbCrLf
feats = feats + 1
Case Is = 5
txtspecial = txtspecial & "Improved Critical" & vbCrLf
ac = ac + 1
Case Is = 4
txtspecial = txtspecial & "Tactical Aid" & vbCrLf
Case Is = 3
txtspecial = txtspecial & "Bonus Feat - Soldier" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 2
txtspecial = txtspecial & "Weapon Specialisation" & vbCrLf
Case Is = 1
txtspecial = txtspecial & "Weapon Focus" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Martial Artist"
ba = ba + b_att(1, Class(8))
will_save = will_save + will(3, Class(8))
ref_save = ref_save + ref(1, Class(8))
fort_save = fort_save + fort(3, Class(8))
hp = hp + (8 + con_mod) * Class(8)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(8)
skill_pts = skill_pts + (2 + int_mod) * Class(8)

Select Case Class(8) 'gets iron fist atacks
Case Is = 10
txtspecial = txtspecial & "Iron Fist - All Attacks" & vbCrLf
Case Is >= 5
txtspecial = txtspecial & "Iron Fist - One Attack" & vbCrLf
End Select

Select Case Class(8) 'gets living weapon
Case Is >= 8
txtspecial = txtspecial & "Living Weapon 1d10" & vbCrLf
Case Is >= 4
txtspecial = txtspecial & "Living Weapon 1d8" & vbCrLf
Case Is >= 1
txtspecial = txtspecial & "Living Weapon 1d6" & vbCrLf
End Select

spec_loop = Class(8)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
ac = ac + 1
Case Is = 9
txtspecial = txtspecial & "Bonus Feat - Martial Artist" & vbCrLf
feats = feats + 1
Case Is = 8
ac = ac + 1
Case Is = 7
txtspecial = txtspecial & "Flurry of Blows" & vbCrLf
ac = ac + 1
Case Is = 6
txtspecial = txtspecial & "Bonus Feat - Martial Artist" & vbCrLf
feats = feats + 1
Case Is = 5
ac = ac + 1
Case Is = 4
ac = ac + 1
Case Is = 3
txtspecial = txtspecial & "Bonus Feat - Martial Artist" & vbCrLf
feats = feats + 1
Case Is = 2
txtspecial = txtspecial & "Flying Kick" & vbCrLf
ac = ac + 1
Case Is = 1
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Gunslinger"
ba = ba + b_att(2, Class(9))
will_save = will_save + will(2, Class(9))
ref_save = ref_save + ref(2, Class(9))
fort_save = fort_save + fort(3, Class(9))
hp = hp + (10 + con_mod) * Class(9)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(9)
skill_pts = skill_pts + (4 + int_mod) * Class(9)

spec_loop = Class(9)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Bullseye" & vbCrLf
Case Is = 9
txtspecial = txtspecial & "Bonus Feat - Gunslinger" & vbCrLf
ac = ac + 1
feats = feats + 1
Case Is = 8
txtspecial = txtspecial & "Greater Weapon Focus" & vbCrLf
Case Is = 7
txtspecial = txtspecial & "Sharp Shooting" & vbCrLf
ac = ac + 1
Case Is = 6
txtspecial = txtspecial & "Bonus Feat - Gunslinger" & vbCrLf
feats = feats + 1
Case Is = 5
txtspecial = txtspecial & "Lightning Shot" & vbCrLf
ac = ac + 1
Case Is = 4
txtspecial = txtspecial & "Defensive Position" & vbCrLf
Case Is = 3
txtspecial = txtspecial & "Bonus Feat - Gunslinger" & vbCrLf
ac = ac + 1
feats = feats + 1
Case Is = 2
txtspecial = txtspecial & "Weapon Focus" & vbCrLf
Case Is = 1
txtspecial = txtspecial & "Close Combat Shot" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Infiltrator"
ba = ba + b_att(3, Class(10))
will_save = will_save + will(3, Class(10))
ref_save = ref_save + ref(1, Class(10))
fort_save = fort_save + fort(3, Class(10))
hp = hp + (8 + con_mod) * Class(10)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(10)
skill_pts = skill_pts + (6 + int_mod) * Class(10)

spec_loop = Class(10)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Without a Trace" & vbCrLf
ac = ac + 1
Case Is = 9
txtspecial = txtspecial & "Bonus Feat - Infiltrator" & vbCrLf
feats = feats + 1
Case Is = 8
txtspecial = txtspecial & "Improved Sweep" & vbCrLf
ac = ac + 1
Case Is = 7
txtspecial = txtspecial & "Improvised Weapon Damage" & vbCrLf
ac = ac + 1
Case Is = 6
txtspecial = txtspecial & "Bonus Feat - Infiltrator" & vbCrLf
feats = feats + 1
Case Is = 5
txtspecial = txtspecial & "Skill Mastery" & vbCrLf
ac = ac + 1
Case Is = 4
txtspecial = txtspecial & "Improved Evasion" & vbCrLf
ac = ac + 1
Case Is = 3
txtspecial = txtspecial & "Bonus Feat - Infiltrator" & vbCrLf
feats = feats + 1
Case Is = 2
txtspecial = txtspecial & "Improvised Implements" & vbCrLf
ac = ac + 1
Case Is = 1
txtspecial = txtspecial & "Sweep" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Daredevil"
ba = ba + b_att(3, Class(11))
will_save = will_save + will(3, Class(11))
ref_save = ref_save + ref(3, Class(11))
fort_save = fort_save + fort(1, Class(11))
hp = hp + (10 + con_mod) * Class(11)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(11)
skill_pts = skill_pts + (4 + int_mod) * Class(11)

Select Case Class(11)
Case Is >= 8
txtspecial = txtspecial & "Adrenaline Rush - Two Ability Scores" & vbCrLf
Case Is >= 5
txtspecial = txtspecial & "Adrenaline Rush - One Ability Scores" & vbCrLf
End Select

spec_loop = Class(11)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Damage Threshold" & vbCrLf
ac = ac + 1
Case Is = 9
txtspecial = txtspecial & "Bonus Feat - Daredevil" & vbCrLf
feats = feats + 1
Case Is = 8
ac = ac + 1
Case Is = 7
txtspecial = txtspecial & "Delay Damage" & vbCrLf
ac = ac + 1
Case Is = 6
txtspecial = txtspecial & "Bonus Feat - Daredevil" & vbCrLf
feats = feats + 1
Case Is = 5
ac = ac + 1
Case Is = 4
txtspecial = txtspecial & "Action Boost" & vbCrLf
ac = ac + 1
Case Is = 3
txtspecial = txtspecial & "Bonus Feat - Daredevil" & vbCrLf
feats = feats + 1
Case Is = 2
txtspecial = txtspecial & "Nip Up" & vbCrLf
ac = ac + 1
Case Is = 1
txtspecial = txtspecial & "Fearless" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Bodyguard"
ba = ba + b_att(2, Class(12))
will_save = will_save + will(3, Class(12))
ref_save = ref_save + ref(1, Class(12))
fort_save = fort_save + fort(2, Class(12))
hp = hp + (12 + con_mod) * Class(12)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(12)
skill_pts = skill_pts + (2 + int_mod) * Class(12)

Select Case Class(12)
Case Is >= 8
txtspecial = txtspecial & "Combat Sense +2" & vbCrLf
Case Is >= 2
txtspecial = txtspecial & "Combat Sense +1" & vbCrLf
End Select

spec_loop = Class(12)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Blanket Protection" & vbCrLf
Case Is = 9
txtspecial = txtspecial & "Bonus Feat - Bodyguard" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 7
txtspecial = txtspecial & "Defensive Strike" & vbCrLf
ac = ac + 1
Case Is = 6
txtspecial = txtspecial & "Bonus Feat - Bodyguard" & vbCrLf
feats = feats + 1
Case Is = 5
txtspecial = txtspecial & "Improved Charge" & vbCrLf
ac = ac + 1
Case Is = 4
txtspecial = txtspecial & "Sudden Action" & vbCrLf
Case Is = 3
txtspecial = txtspecial & "Bonus Feat - Bodyguard" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 1
txtspecial = txtspecial & "Harm's Way" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Field Scientist"
ba = ba + b_att(3, Class(13))
will_save = will_save + will(3, Class(13))
ref_save = ref_save + ref(2, Class(13))
fort_save = fort_save + fort(2, Class(13))
hp = hp + 8 * Class(13)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(13)
skill_pts = skill_pts + (6 + int_mod) * Class(13)

spec_loop = Class(13)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Major Breakthrough" & vbCrLf
Case Is = 9
txtspecial = txtspecial & "Bonus Feat - Field Scientist" & vbCrLf
feats = feats + 1
Case Is = 8
txtspecial = txtspecial & "Smart Weapon" & vbCrLf
ac = ac + 1
Case Is = 7
txtspecial = txtspecial & "Smart Survival" & vbCrLf
Case Is = 6
txtspecial = txtspecial & "Bonus Feat - Field Scientist" & vbCrLf
feats = feats + 1
Case Is = 5
txtspecial = txtspecial & "Minor Breakthrough" & vbCrLf
ac = ac + 1
Case Is = 4
txtspecial = txtspecial & "Skill Mastery" & vbCrLf
Case Is = 3
txtspecial = txtspecial & "Bonus Feat - Field Scientist" & vbCrLf
feats = feats + 1
Case Is = 2
txtspecial = txtspecial & "Scientific Improvisation" & vbCrLf
ac = ac + 1
Case Is = 1
txtspecial = txtspecial & "Smart Defense" & vbCrLf
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Techie"
ba = ba + b_att(3, Class(14))
will_save = will_save + will(1, Class(14))
ref_save = ref_save + ref(3, Class(14))
fort_save = fort_save + fort(3, Class(14))
hp = hp + (6 + con_mod) * Class(14)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(14)
skill_pts = skill_pts + (6 + int_mod) * Class(14)

Select Case Class(14)
Case Is >= 1
txtspecial = txtspecial & "Jury Rig +4" & vbCrLf
Case Is >= 7
txtspecial = txtspecial & "Jury Rig +2" & vbCrLf
End Select

spec_loop = Class(14)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Mastercraft" & vbCrLf
Case Is = 9
txtspecial = txtspecial & "Bonus Feat - Techie" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 8
txtspecial = txtspecial & "Mastercraft" & vbCrLf
Case Is = 7
ac = ac + 1
Case Is = 6
txtspecial = txtspecial & "Bonus Feat - Techie" & vbCrLf
feats = feats + 1
Case Is = 5
txtspecial = txtspecial & "Mastercraft" & vbCrLf
ac = ac + 1
Case Is = 4
txtspecial = txtspecial & "Build Robot" & vbCrLf
Case Is = 3
txtspecial = txtspecial & "Bonus Feat - Techie" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 2
txtspecial = txtspecial & "Extreme Machine" & vbCrLf
Case Is = 1
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Field Medic"
ba = ba + b_att(3, Class(15))
will_save = will_save + will(2, Class(15))
ref_save = ref_save + ref(3, Class(15))
fort_save = fort_save + fort(1, Class(15))
hp = hp + (8 + con_mod) * Class(15)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(15)
skill_pts = skill_pts + (4 + int_mod) * Class(15)

Select Case Class(15)
Case Is >= 8
txtspecial = txtspecial & "Medical Specialist +3" & vbCrLf
Case Is >= 5
txtspecial = txtspecial & "Medical Specialist +2" & vbCrLf
Case Is >= 1
txtspecial = txtspecial & "Medical Specialist +1" & vbCrLf
End Select

spec_loop = Class(15)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Medical Miracle" & vbCrLf
Case Is = 9
txtspecial = txtspecial & "Bonus Feat - Field Medic" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 7
txtspecial = txtspecial & "Minor Medical Miracle" & vbCrLf
ac = ac + 1
Case Is = 6
txtspecial = txtspecial & "Bonus Feat - Field Medic" & vbCrLf
feats = feats + 1
Case Is = 5
ac = ac + 1
Case Is = 4
txtspecial = txtspecial & "Medical Mastery" & vbCrLf
Case Is = 3
txtspecial = txtspecial & "Bonus Feat - Field Medic" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 2
txtspecial = txtspecial & "Expert Healer" & vbCrLf
Case Is = 1
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Investigator"
ba = ba + b_att(2, Class(16))
will_save = will_save + will(2, Class(16))
ref_save = ref_save + ref(2, Class(16))
fort_save = fort_save + fort(3, Class(16))
hp = hp + (6 + con_mod) * Class(16)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(16)
skill_pts = skill_pts + (4 + int_mod) * Class(16)

spec_loop = Class(16)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Sixth Sense" & vbCrLf
Case Is = 9
txtspecial = txtspecial & "Bonus Feat - Investigator" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 8
txtspecial = txtspecial & "Contact - High Level" & vbCrLf
Case Is = 7
txtspecial = txtspecial & "Discern Lie" & vbCrLf
ac = ac + 1
Case Is = 6
txtspecial = txtspecial & "Bonus Feat - Investigator" & vbCrLf
feats = feats + 1
Case Is = 5
txtspecial = txtspecial & "Contact - Mid Level" & vbCrLf
ac = ac + 1
Case Is = 4
txtspecial = txtspecial & "Non-Lethal Force" & vbCrLf
Case Is = 3
txtspecial = txtspecial & "Bonus Feat - Investigator" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 2
txtspecial = txtspecial & "Contact - Low Level" & vbCrLf
Case Is = 1
txtspecial = txtspecial & "Profile" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Personality"
ba = ba + b_att(3, Class(17))
will_save = will_save + will(3, Class(17))
ref_save = ref_save + ref(2, Class(17))
fort_save = fort_save + fort(2, Class(17))
hp = hp + (6 + con_mod) * Class(17)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(17)
skill_pts = skill_pts + (4 + int_mod) * Class(17)

spec_loop = Class(17)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Compelling Performance" & vbCrLf
Case Is = 9
txtspecial = txtspecial & "Bonus Feat - Personality" & vbCrLf
feats = feats + 1
Case Is = 8
txtspecial = txtspecial & "Royalty" & vbCrLf
ac = ac + 1
Case Is = 7
txtspecial = txtspecial & "Bonus Class Skill" & vbCrLf
Case Is = 6
txtspecial = txtspecial & "Bonus Feat - Personality" & vbCrLf
feats = feats + 1
Case Is = 5
txtspecial = txtspecial & "winning Smile" & vbCrLf
ac = ac + 1
Case Is = 4
txtspecial = txtspecial & "Royalty" & vbCrLf
Case Is = 3
txtspecial = txtspecial & "Bonus Feat - Personality" & vbCrLf
feats = feats + 1
Case Is = 2
txtspecial = txtspecial & "Bonus Class Skill" & vbCrLf
ac = ac + 1
Case Is = 1
txtspecial = txtspecial & "Unlimited Access" & vbCrLf
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Negotiator"
ba = ba + b_att(2, Class(18))
will_save = will_save + will(1, Class(18))
ref_save = ref_save + ref(3, Class(18))
fort_save = fort_save + fort(2, Class(18))
hp = hp + (8 + con_mod) * Class(18)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(18)
skill_pts = skill_pts + (4 + int_mod) * Class(18)

Select Case Class(18)
Case Is = 10
txtspecial = txtspecial & "Talk Down - All Opponents" & vbCrLf
Case Is >= 7
txtspecial = txtspecial & "Talk Down - Several Opponents" & vbCrLf
Case Is >= 4
txtspecial = txtspecial & "Talk Down - One Opponents" & vbCrLf
End Select

spec_loop = Class(18)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 9
txtspecial = txtspecial & "Bonus Feat - Negotiator" & vbCrLf
feats = feats + 1
Case Is = 8
txtspecial = txtspecial & "Sow Distrust" & vbCrLf
ac = ac + 1
Case Is = 6
txtspecial = txtspecial & "Bonus Feat - Negotiator" & vbCrLf
feats = feats + 1
Case Is = 5
txtspecial = txtspecial & "No Sweat" & vbCrLf
ac = ac + 1
Case Is = 3
txtspecial = txtspecial & "Bonus Feat - Negotiator" & vbCrLf
feats = feats + 1
Case Is = 2
txtspecial = txtspecial & "React First" & vbCrLf
ac = ac + 1
Case Is = 1
txtspecial = txtspecial & "Conceal Motive" & vbCrLf
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0

End Select

Next increment

counter(3) = counter(3) + 1

Loop Until select_cls = ""

txtmods = "BA" & vbTab & ba & vbCrLf & "HP" & vbTab & hp & vbCrLf & "WILL" & vbTab & will_save & vbCrLf & "FORT" & vbTab & fort_save & vbCrLf & "REF" & vbTab & ref_save
frmcharsheet.txtmisc = frmcharsheet.txtmisc & "Classes" & vbTab & vbTab & clslvls
End Sub

Sub naruto_classes(ByVal level As Integer, ByRef ba As Integer, ByRef b_att() As Integer, ByRef will_save As Integer, ByRef will() As Integer, ByRef ref_save As Integer, ByRef ref() As Integer, ByRef fort_save As Integer, ByRef fort() As Integer, ByRef hp As Integer, ByVal con_mod As Integer, ByRef skill_pts As Integer, ByVal int_mod As Integer, ByRef ac As Integer, ByRef feats As Integer)

Dim counter(3) As Integer
Dim Class(35) As Integer 'finds out the levels in each class
Dim lvl As Integer 'individual level in class
Dim ttl_lvl As Integer 'total levels used
Dim cls As String
Dim chosen(35) As String
Dim selected As Integer
Dim select_cls As String
Dim increment As Integer
Dim first_class As String

ttl_lvl = 0
counter(2) = 1
counter(3) = 1

Do
For counter(1) = 1 To level
If ttl_lvl < level Then

If ttl_lvl < 3 Then
Do
cls = InputBox("Select the class that you wish to take. Initial letter must be uppercase, i.e. Strong, Fast. Can not be prestige.")
Loop Until cls = "Strong" Or cls = "Fast" Or cls = "Tough" Or cls = "Smart" Or cls = "Dedicated" Or cls = "Charismatic"
Else

Do
cls = InputBox("Select the class that you wish to take. Initial letter must be uppercase, i.e. Strong, Fast, Technique Analyst.")
Loop Until cls = "Strong" Or cls = "Fast" Or cls = "Tough" Or cls = "Smart" Or cls = "Dedicated" Or cls = "Charismatic" Or cls = "Beast Lord" Or cls = "Beast Master" Or cls = "Blink Strike" Or cls = "Devastator" Or cls = "Elementalist" Or cls = "Exarch" Or cls = "Exemplar" Or cls = "Genjutsu Master" Or cls = "Livewire" Or cls = "Master Strategist" Or cls = "Medical Specialist" Or cls = "Ninja Operations Counter" Or cls = "Ninja Police" Or cls = "Ninja Scout" Or cls = "Puppeteer" Or cls = "Sacred Fist" Or cls = "Samurai" Or cls = "Shade" Or cls = "Shinobi Adept" Or cls = "Shinobi Bodyguard" Or cls = "Shinobi Swordsman" Or cls = "Shuriken Expert" Or cls = "Soul Edge" Or cls = "Squad Captain" Or cls = "Summoner" Or cls = "Sword Savant" Or cls = "Taijutsu Master" Or cls = "Technique Analyst" Or cls = "Weapon Master"
End If

If counter(1) = 1 Then
first_class = cls
End If

chosen(counter(2)) = cls

Do
lvl = InputBox("Select how many levels you wish to take in that class. You have " & level - ttl_lvl & " levels left to spend.")

'makes sure that you can't take over the maximum amount of levels in a class

If cls = "Strong" Or cls = "Fast" Or cls = "Tough" Or cls = "Smart" Or cls = "Dedicated" Or cls = "Charismatic" Or cls = "Beast Master" Or cls = "Medical Specialist" Or cls = "Ninja Police" Or cls = "Ninja Scout" Or cls = "Puppeteer" Or cls = "Sacred Fist" Or cls = "Samurai" Or cls = "Shinobi Bodyguard" Or cls = "Shinobi Swordsman" Or cls = "Shuriken Expert" Or cls = "Soul Edge" Or cls = "Squad Captain" Or cls = "Taijutsu Master" Then

Do

If lvl > 10 Or lvl < 0 Then
lvl = InputBox("Max of 10 levels in class you wish to take in that class.")
End If
Loop Until lvl <= 10
End If


If cls = "Elementalist" Or cls = "Genjutsu Master" Or cls = "Master Strategist" Or cls = "Summoner" Or cls = "Sword Savant" Then

Do

If lvl > 7 Or lvl < 0 Then
lvl = InputBox("Max of 7 levels in class you wish to take in that class.")
End If
Loop Until lvl <= 7
End If


If cls = "Blink Strike" Or cls = "Devastator" Or cls = "Exarch" Or cls = "Exemplar" Or cls = "Ninja Operations Counter" Or cls = "Shade" Or cls = "Technique Analyst" Or cls = "Weapon Master" Then

Do

If lvl > 5 Or lvl < 0 Then
lvl = InputBox("Max of 5 levels in class you wish to take in that class.")
End If
Loop Until lvl <= 5
End If


If cls = "Beast Lord" Or cls = "Livewire" Or cls = "Shinobi Adept" Then

Do

If lvl > 3 Or lvl < 0 Then
lvl = InputBox("Max of 3 levels in class you wish to take in that class.")
End If
Loop Until lvl <= 3
End If

Loop Until lvl <= level 'the user can't continue if he has more than his maximum level
ttl_lvl = ttl_lvl + lvl



Select Case cls 'finds out how many levels are in each class
Case Is = "Strong"
Class(1) = Class(1) + lvl
If Class(1) > 10 Then
Class(1) = Class(1) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Fast"
Class(2) = Class(2) + lvl
If Class(2) > 10 Then
Class(2) = Class(2) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Tough"
Class(3) = Class(3) + lvl
If Class(3) > 10 Then
Class(3) = Class(3) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Smart"
Class(4) = Class(4) + lvl
If Class(4) > 10 Then
Class(4) = Class(4) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Dedicated"
Class(5) = Class(5) + lvl
If Class(5) > 10 Then
Class(5) = Class(5) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Charismatic"
Class(6) = Class(6) + lvl
If Class(6) > 10 Then
Class(6) = Class(6) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Beast Lord"
Class(7) = Class(7) + lvl
If Class(7) > 3 Then
Class(7) = Class(7) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Beast Master"
Class(8) = Class(8) + lvl
If Class(8) > 10 Then
Class(8) = Class(8) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Blink Strike"
Class(9) = Class(9) + lvl
If Class(9) > 5 Then
Class(9) = Class(9) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Devastator"
Class(10) = Class(10) + lvl
If Class(10) > 5 Then
Class(10) = Class(10) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Elementalist"
Class(11) = Class(11) + lvl
If Class(11) > 7 Then
Class(11) = Class(11) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Exarch"
Class(12) = Class(12) + lvl
If Class(12) > 5 Then
Class(12) = Class(12) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Exemplar"
Class(13) = Class(13) + lvl
If Class(13) > 5 Then
Class(13) = Class(13) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Genjutsu Master"
Class(14) = Class(14) + lvl
If Class(14) > 7 Then
Class(14) = Class(14) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Livewire"
Class(15) = Class(15) + lvl
If Class(15) > 3 Then
Class(15) = Class(15) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Master Strategist"
Class(16) = Class(16) + lvl
If Class(16) > 7 Then
Class(16) = Class(16) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Medical Specialist"
Class(17) = Class(17) + lvl
If Class(17) > 10 Then
Class(17) = Class(17) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Ninja Operations Counter"
Class(18) = Class(18) + lvl
If Class(18) > 5 Then
Class(18) = Class(18) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Ninja Police"
Class(19) = Class(19) + lvl
If Class(19) > 10 Then
Class(19) = Class(19) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Ninja Scout"
Class(20) = Class(20) + lvl
If Class(20) > 10 Then
Class(20) = Class(20) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Puppeteer"
Class(21) = Class(21) + lvl
If Class(21) > 10 Then
Class(21) = Class(21) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Sacred Fist"
Class(22) = Class(22) + lvl
If Class(22) > 10 Then
Class(22) = Class(22) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Samurai"
Class(23) = Class(23) + lvl
If Class(23) > 10 Then
Class(23) = Class(23) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Shade"
Class(24) = Class(24) + lvl
If Class(24) > 5 Then
Class(24) = Class(24) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Shinobi Adept"
Class(25) = Class(25) + lvl
If Class(25) > 3 Then
Class(25) = Class(25) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Shinobi Bodyguard"
Class(26) = Class(26) + lvl
If Class(26) > 10 Then
Class(26) = Class(26) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Shinobi Swordsman"
Class(27) = Class(27) + lvl
If Class(27) > 10 Then
Class(27) = Class(27) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Shuriken Expert"
Class(28) = Class(28) + lvl
If Class(28) > 10 Then
Class(28) = Class(28) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Soul Edge"
Class(29) = Class(29) + lvl
If Class(29) > 10 Then
Class(29) = Class(29) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Squad Captain"
Class(30) = Class(30) + lvl
If Class(30) > 10 Then
Class(30) = Class(30) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Summoner"
Class(31) = Class(31) + lvl
If Class(31) > 7 Then
Class(31) = Class(31) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Sword Savant"
Class(32) = Class(32) + lvl
If Class(32) > 7 Then
Class(32) = Class(32) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Taijutsu Master"
Class(33) = Class(33) + lvl
If Class(33) > 10 Then
Class(33) = Class(33) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Technique Analyst"
Class(34) = Class(34) + lvl
If Class(34) > 5 Then
Class(34) = Class(34) - lvl
MsgBox "Class can not have that many levels in it."
End If

Case Is = "Weapon Master"
Class(35) = Class(35) + lvl
If Class(35) > 5 Then
Class(35) = Class(35) - lvl
MsgBox "Class can not have that many levels in it."
End If

End Select


counter(2) = counter(2) + 1

End If

Next counter(1)

Loop Until ttl_lvl = level


increment = 1

Do 'makes sure that no duplicates appear in chosen class

select_cls = chosen(increment)

For increment = 1 To counter(2)
selected = increment + 1
If select_cls = chosen(selected) Then
chosen(selected) = "x"
End If
Next increment

Dim spec_loop As Integer 'gets the specials from classes
Dim clslvls As String 'puts down how many levels are in each class, which will be put into the misc area afterwards

increment = 1

For increment = 1 To counter(2)
select_cls = chosen(increment)

Select Case select_cls

Case Is = "Strong"
ba = ba + b_att(1, Class(1))
will_save = will_save + will(3, Class(1))
ref_save = ref_save + ref(3, Class(1))
fort_save = fort_save + fort(2, Class(1))
hp = hp + (8 + con_mod) * Class(1)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(1)

If first_class = select_cls Then
skill_pts = skill_pts + (2 + int_mod) * (Class(1) + 3)
Else
skill_pts = skill_pts + (2 + int_mod) * Class(1)
End If

spec_loop = Class(1)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Bonus Feat - Strong" & vbCrLf
feats = feats + 1
Case Is = 9
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
ac = ac + 1
Case Is = 8
txtspecial = txtspecial & "Bonus Feat - Strong" & vbCrLf
feats = feats + 1
Case Is = 7
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
ac = ac + 1
Case Is = 6
txtspecial = txtspecial & "Bonus Feat - Strong" & vbCrLf
feats = feats + 1
Case Is = 5
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
Case Is = 4
txtspecial = txtspecial & "Bonus Feat - Strong" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 3
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
Case Is = 2
txtspecial = txtspecial & "Bonus Feat - Strong" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 1
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Fast"
ba = ba + b_att(2, Class(2))
will_save = will_save + will(3, Class(2))
ref_save = ref_save + ref(2, Class(2))
fort_save = fort_save + fort(3, Class(2))
hp = hp + (8 + con_mod) * Class(2)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(2)

If first_class = select_cls Then
skill_pts = skill_pts + (4 + int_mod) * (Class(2) + 3)
Else
skill_pts = skill_pts + (4 + int_mod) * Class(2)
End If

spec_loop = Class(2)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Bonus Feat - Fast" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 9
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
Case Is = 8
txtspecial = txtspecial & "Bonus Feat - Fast" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 7
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
Case Is = 6
txtspecial = txtspecial & "Bonus Feat - Fast" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 5
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
Case Is = 4
txtspecial = txtspecial & "Bonus Feat - Fast" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 3
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
Case Is = 2
txtspecial = txtspecial & "Bonus Feat - Fast" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 1
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
ac = ac + 3
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Tough"
ba = ba + b_att(2, Class(3))
will_save = will_save + will(3, Class(3))
ref_save = ref_save + ref(3, Class(3))
fort_save = fort_save + fort(2, Class(3))
hp = hp + (10 + con_mod) * Class(3)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(3)

If first_class = select_cls Then
skill_pts = skill_pts + (2 + int_mod) * (Class(3) + 3)
Else
skill_pts = skill_pts + (2 + int_mod) * Class(3)
End If

spec_loop = Class(3)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Bonus Feat - Tough" & vbCrLf
feats = feats + 1
Case Is = 9
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
ac = ac + 1
Case Is = 8
txtspecial = txtspecial & "Bonus Feat - Tough" & vbCrLf
feats = feats + 1
Case Is = 7
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
ac = ac + 1
Case Is = 6
txtspecial = txtspecial & "Bonus Feat - Tough" & vbCrLf
feats = feats + 1
Case Is = 5
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
Case Is = 4
txtspecial = txtspecial & "Bonus Feat - Tough" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 3
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
Case Is = 2
txtspecial = txtspecial & "Bonus Feat - Tough" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 1
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Smart"
ba = ba + b_att(3, Class(4))
will_save = will_save + will(2, Class(4))
ref_save = ref_save + ref(3, Class(4))
fort_save = fort_save + fort(3, Class(4))
hp = hp + (6 + con_mod) * Class(4)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(4)

If first_class = select_cls Then
skill_pts = skill_pts + (8 + int_mod) * (Class(4) + 3)
Else
skill_pts = skill_pts + (8 + int_mod) * Class(4)
End If

spec_loop = Class(4)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Bonus Feat - Smart" & vbCrLf
feats = feats + 1
Case Is = 9
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
Case Is = 8
txtspecial = txtspecial & "Bonus Feat - Smart" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 7
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
Case Is = 6
txtspecial = txtspecial & "Bonus Feat - Smart" & vbCrLf
feats = feats + 1
Case Is = 5
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
ac = ac + 1
Case Is = 4
txtspecial = txtspecial & "Bonus Feat - Smart" & vbCrLf
feats = feats + 1
Case Is = 3
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
Case Is = 2
txtspecial = txtspecial & "Bonus Feat - Smart" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 1
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Dedicated"
ba = ba + b_att(2, Class(5))
will_save = will_save + will(2, Class(5))
ref_save = ref_save + ref(3, Class(5))
fort_save = fort_save + fort(2, Class(5))
hp = hp + (6 + con_mod) * Class(5)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(5)

If first_class = select_cls Then
skill_pts = skill_pts + (4 + int_mod) * (Class(5) + 3)
Else
skill_pts = skill_pts + (4 + int_mod) * Class(5)
End If

spec_loop = Class(5)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Bonus Feat - Dedicated" & vbCrLf
feats = feats + 1
Case Is = 9
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
ac = ac + 1
Case Is = 8
txtspecial = txtspecial & "Bonus Feat - Dedicated" & vbCrLf
feats = feats + 1
Case Is = 7
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
ac = ac + 1
Case Is = 6
txtspecial = txtspecial & "Bonus Feat - Dedicated" & vbCrLf
feats = feats + 1
Case Is = 5
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
Case Is = 4
txtspecial = txtspecial & "Bonus Feat - Dedicated" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 3
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
Case Is = 2
txtspecial = txtspecial & "Bonus Feat - Dedicated" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 1
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Charismatic"
ba = ba + b_att(3, Class(6))
will_save = will_save + will(3, Class(6))
ref_save = ref_save + ref(2, Class(6))
fort_save = fort_save + fort(2, Class(6))
hp = hp + (6 + con_mod) * Class(6)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(6)

If first_class = select_cls Then
skill_pts = skill_pts + (6 + int_mod) * (Class(6) + 3)
Else
skill_pts = skill_pts + (6 + int_mod) * Class(6)
End If

spec_loop = Class(6)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Bonus Feat - Charismatic" & vbCrLf
feats = feats + 1
Case Is = 9
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
Case Is = 8
txtspecial = txtspecial & "Bonus Feat - Charismatic" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 7
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
Case Is = 6
txtspecial = txtspecial & "Bonus Feat - Charismatic" & vbCrLf
feats = feats + 1
Case Is = 5
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
ac = ac + 1
Case Is = 4
txtspecial = txtspecial & "Bonus Feat - Charismatic" & vbCrLf
feats = feats + 1
Case Is = 3
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
Case Is = 2
txtspecial = txtspecial & "Bonus Feat - Charismatic" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 1
txtspecial = txtspecial & select_cls & " - Talent" & vbCrLf
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Beastlord"
ba = ba + b_att(1, Class(7))
will_save = will_save + will(3, Class(7))
ref_save = ref_save + ref(1, Class(7))
fort_save = fort_save + fort(1, Class(7))
hp = hp + (10 + con_mod) * Class(7)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(7)
skill_pts = skill_pts + (2 + int_mod) * Class(7)

spec_loop = Class(7)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop

Case Is = 3
txtspecial = txtspecial & "Extra Animal Companion -6" & vbCrLf & "Lowlight Vision" & vbCrLf
ac = ac + 1
Case Is = 2
txtspecial = txtspecial & "Aspect of the Pack" & vbCrLf
Case Is = 1
txtspecial = txtspecial & "Extra Animal Companion -3" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Beastmaster"
ba = ba + b_att(1, Class(8))
will_save = will_save + will(3, Class(8))
ref_save = ref_save + ref(2, Class(8))
fort_save = fort_save + fort(1, Class(8))
hp = hp + (10 + con_mod) * Class(8)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(8)
skill_pts = skill_pts + (2 + int_mod) * Class(8)

Select Case Class(8) 'gets frenzy/day
Case Is >= 9
txtspecial = txtspecial & "Frenzy 3/Day" & vbCrLf
Case Is >= 5
txtspecial = txtspecial & "Frenzy 2/Day" & vbCrLf
Case Is >= 1
txtspecial = txtspecial & "Frenzy 2/Day" & vbCrLf
End Select

Select Case Class(8) 'feral combat 1dx
Case Is = 10
txtspecial = txtspecial & "Feral Combat 1d10" & vbCrLf
Case Is >= 6
txtspecial = txtspecial & "Feral Combat 1d8" & vbCrLf
Case Is >= 2
txtspecial = txtspecial & "Feral Combat 1d6" & vbCrLf
End Select

Select Case Class(8) 'amazing tricks
Case Is >= 8
txtspecial = txtspecial & "Amazing Tricks +4" & vbCrLf
Case Is >= 4
txtspecial = txtspecial & "Amazing Tricks +2" & vbCrLf
End Select

spec_loop = Class(8)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Animal Aspect" & vbCrLf
Case Is = 9
ac = ac + 1
Case Is = 7
txtspecial = txtspecial & "Greater Frenzy" & vbCrLf
ac = ac + 1
Case Is = 5
ac = ac + 1
Case Is = 3
txtspecial = txtspecial & "Speak with Animals" & vbCrLf
ac = ac + 1
Case Is = 2
txtspecial = txtspecial & "Inspire Frenzy" & vbCrLf
Case Is = 1
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Blinkstrike"
ba = ba + b_att(1, Class(9))
will_save = will_save + will(3, Class(9))
ref_save = ref_save + ref(1, Class(9))
fort_save = fort_save + fort(3, Class(9))
hp = hp + (8 + con_mod) * Class(9)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(9)
skill_pts = skill_pts + (2 + int_mod) * Class(9)

Select Case Class(9) 'blinkstrike +x
Case Is = 5
txtspecial = txtspecial & "Blinkstrike +5" & vbCrLf
Case Is >= 3
txtspecial = txtspecial & "Blinkstrike +3" & vbCrLf
Case Is >= 1
txtspecial = txtspecial & "Blinkstrike +1" & vbCrLf
End Select

Select Case Class(9) 'blink step/day
Case Is >= 4
txtspecial = txtspecial & "Blink Step 4/Day" & vbCrLf
Case Is >= 2
txtspecial = txtspecial & "Blink Step 2/Day" & vbCrLf
End Select

spec_loop = Class(9)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 5
txtspecial = txtspecial & "Warp Charge" & vbCrLf
ac = ac + 1
Case Is = 4
txtspecial = txtspecial & "Evasion X" & vbCrLf
ac = ac + 1
Case Is = 2
txtspecial = txtspecial & "Weapon Focus" & vbCrLf
ac = ac + 1
Case Is = 1
txtspecial = txtspecial & "Bonus Chakra" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Devastator"
ba = ba + b_att(2, Class(10))
will_save = will_save + will(1, Class(10))
ref_save = ref_save + ref(2, Class(10))
fort_save = fort_save + fort(3, Class(10))
hp = hp + (6 + con_mod) * Class(10)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(10)
skill_pts = skill_pts + (4 + int_mod) * Class(10)

spec_loop = Class(10)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 5
txtspecial = txtspecial & "Force of nature" & vbCrLf
Case Is = 4
txtspecial = txtspecial & "Succession Technique" & vbCrLf
ac = ac + 1
Case Is = 3
txtspecial = txtspecial & "Succession Technique" & vbCrLf
Case Is = 2
txtspecial = txtspecial & "Succession Technique" & vbCrLf
ac = ac + 1
Case Is = 1
txtspecial = txtspecial & "Unleashed Power" & vbCrLf & "Bonus Chakra" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Elementalist"
ba = ba + b_att(2, Class(11))
will_save = will_save + will(1, Class(11))
ref_save = ref_save + ref(2, Class(11))
fort_save = fort_save + fort(3, Class(11))
hp = hp + (6 + con_mod) * Class(11)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(11)
skill_pts = skill_pts + (4 + int_mod) * Class(11)

spec_loop = Class(11)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 7
txtspecial = txtspecial & "Elemental Surge" & vbCrLf
ac = ac + 1
Case Is = 6
txtspecial = txtspecial & "Elemental Focus" & vbCrLf
Case Is = 5
txtspecial = txtspecial & "Rage of the Elements" & vbCrLf
ac = ac + 1
Case Is = 4
txtspecial = txtspecial & "Bonus Feat - Elementalist" & vbCrLf
feats = feats + 1
Case Is = 3
txtspecial = txtspecial & "Limitless Fury" & vbCrLf
ac = ac + 1
Case Is = 2
txtspecial = txtspecial & "Elemental Fury" & vbCrLf
Case Is = 1
txtspecial = txtspecial & "Elemental Specialisation" & vbCrLf & "Bonus Chakra" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Exarch"
ba = ba + b_att(3, Class(12))
will_save = will_save + will(1, Class(12))
ref_save = ref_save + ref(3, Class(12))
fort_save = fort_save + fort(1, Class(12))
hp = hp + (6 + con_mod) * Class(12)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(12)
skill_pts = skill_pts + (4 + int_mod) * Class(12)

spec_loop = Class(12)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 5
txtspecial = txtspecial & "Exarch's Blessing" & vbCrLf
Case Is = 4
txtspecial = txtspecial & "Exarch Arcana" & vbCrLf
ac = ac + 1
Case Is = 3
txtspecial = txtspecial & "Exarch Arcana" & vbCrLf
Case Is = 2
txtspecial = txtspecial & "Exarch Arcana" & vbCrLf
ac = ac + 1
Case Is = 1
txtspecial = txtspecial & "Chakra Scalpel Overchannel" & vbCrLf & "Bonus Chakra" & vbCrLf & "Medical Specialist Abilities" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Exemplar"
ba = ba + b_att(1, Class(13))
will_save = will_save + will(3, Class(13))
ref_save = ref_save + ref(2, Class(13))
fort_save = fort_save + fort(1, Class(13))
hp = hp + 10 * Class(13)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(13)
skill_pts = skill_pts + (2 + int_mod) * Class(13)

spec_loop = Class(13)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 5
txtspecial = txtspecial & "Last Stand" & vbCrLf
ac = ac + 1
Case Is = 4
txtspecial = txtspecial & "High Mastery" & vbCrLf
ac = ac + 1
Case Is = 3
txtspecial = txtspecial & "High Mastery" & vbCrLf
Case Is = 2
txtspecial = txtspecial & "High Mastery" & vbCrLf
ac = ac + 1
Case Is = 1
txtspecial = txtspecial & "Master Strike" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Genjutsu Master"
ba = ba + b_att(2, Class(14))
will_save = will_save + will(1, Class(14))
ref_save = ref_save + ref(3, Class(14))
fort_save = fort_save + fort(2, Class(14))
hp = hp + (6 + con_mod) * Class(14)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(14)
skill_pts = skill_pts + (4 + int_mod) * Class(14)

Select Case Class(14) 'sneak attack
Case Is = 7
txtspecial = txtspecial & "Sneak Attack +1d6" & vbCrLf
Case Is >= 3
txtspecial = txtspecial & "Sneak Attack +2d6" & vbCrLf
End Select

spec_loop = Class(14)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 7
txtspecial = txtspecial & "Genjutsu Mastery" & vbCrLf
Case Is = 6
txtspecial = txtspecial & "Quicken Illusion" & vbCrLf
ac = ac + 1
Case Is = 5
txtspecial = txtspecial & "Genjutsu Mastery" & vbCrLf
Case Is = 4
txtspecial = txtspecial & "Bonus Feat - Genjutsu Master" & vbCrLff
feats = feats + 1
ac = ac + 1
Case Is = 3
txtspecial = txtspecial & "Genjutsu Mastery" & vbCrLf
Case Is = 2
txtspecial = txtspecial & "Bonus Feat - Genjutsu Master" & vbCrLff
feats = feats + 1
ac = ac + 1
Case Is = 1
txtspecial = txtspecial & "Genjutsu Mastery" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Livewire"
ba = ba + b_att(1, Class(15))
will_save = will_save + will(3, Class(15))
ref_save = ref_save + ref(1, Class(15))
fort_save = fort_save + fort(3, Class(15))
hp = hp + (8 + con_mod) * Class(15)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(15)
skill_pts = skill_pts + (2 + int_mod) * Class(15)

spec_loop = Class(15)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 3
txtspecial = txtspecial & "Wire Trick" & vbCrLf & "Livewire" & vbCrLf
Case Is = 2
txtspecial = txtspecial & "Wire Trick" & vbCrLf & "Bonus Feat - Livewire" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 1
txtspecial = txtspecial & "Wire Trick" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Master Strategist"
ba = ba + b_att(3, Class(16))
will_save = will_save + will(1, Class(16))
ref_save = ref_save + ref(3, Class(16))
fort_save = fort_save + fort(2, Class(16))
hp = hp + (6 + con_mod) * Class(16)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(16)
skill_pts = skill_pts + (8 + int_mod) * Class(16)

Select Case Class(16) 'swift planning/day
Case Is >= 5
txtspecial = txtspecial & "Swift Planning 2/Day" & vbCrLf
Case Is >= 2
txtspecial = txtspecial & "Swift Planning 1/Day" & vbCrLf
End Select

spec_loop = Class(16)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 7
txtspecial = txtspecial & "Checkmate" & vbCrLf
ac = ac + 1
Case Is = 6
txtspecial = txtspecial & "Tactical Focus" & vbCrLf & "Bonus Feat - Master Strategist" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 4
txtspecial = txtspecial & "Declaration of War" & vbCrLf & "Tactical Assessment" & vbCrLf & "Bonus Feat - Master Strategist" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 2
txtspecial = txtspecial & "Strategic Timing" & vbCrLf & "Bonus Feat - Master Strategist" & vbCrLf
feats = feats + 1
Case Is = 1
txtspecial = txtspecial & "Improved Plans" & vbCrLf & "Bonus Chakra" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Medical Specialist"
ba = ba + b_att(2, Class(17))
will_save = will_save + will(1, Class(17))
ref_save = ref_save + ref(2, Class(17))
fort_save = fort_save + fort(3, Class(17))
hp = hp + (6 + con_mod) * Class(17)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(17)
skill_pts = skill_pts + (4 + int_mod) * Class(17)

Select Case Class(17)
Case Is >= 9
txtspecial = txtspecial & "Chakra Scalpel 1d6" & vbCrLf
Case Is >= 1
txtspecial = txtspecial & "Chakra Scalpel 1d4" & vbCrLf
End Select

Select Case Class(17)
Case Is >= 8
txtspecial = txtspecial & "Sneak Attack + 2d6" & vbCrLf
Case Is >= 3
txtspecial = txtspecial & "Sneak Attack + 1d6" & vbCrLf
End Select

spec_loop = Class(17)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Empower Healing" & vbCrLf
Case Is = 9
ac = ac + 1
Case Is = 7
txtspecial = txtspecial & "Medical Mastery" & vbCrLf
ac = ac + 1
Case Is = 6
txtspecial = txtspecial & "Bonus Feat - Medical Specialist" & vbCrLf
feats = feats + 1
Case Is = 5
txtspecial = txtspecial & "Chakra Scalpel Expertise" & vbCrLf
ac = ac + 1
Case Is = 4
txtspecial = txtspecial & "Bonus Feat - Medical Specialist" & vbCrLf
feats = feats + 1
Case Is = 3
ac = ac + 1
Case Is = 2
txtspecial = txtspecial & "Expert Healer" & vbCrLf & "Bonus Feat - Medical Specialist" & vbCrLf
feats = feats + 1
Case Is = 1
txtspecial = txtspecial & "Medcial Ability" & vbCrLf & "Bonus Chakra" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Ninja Operations Counter"
ba = ba + b_att(2, Class(18))
will_save = will_save + will(2, Class(18))
ref_save = ref_save + ref(2, Class(18))
fort_save = fort_save + fort(2, Class(18))
hp = hp + (6 + con_mod) * Class(18)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(18)
skill_pts = skill_pts + (6 + int_mod) * Class(18)

spec_loop = Class(18)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 5
txtspecial = txtspecial & "Greater Technique Counter" & vbCrLf & "Evasion X" & vbCrLf
ac = ac + 1
Case Is = 4
txtspecial = txtspecial & "Bonus Feat - Ninja Operations Counter" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 3
txtspecial = txtspecial & "Swift Tracker" & vbCrLf & "Tenketsu Freeze" & vbCrLf
Case Is = 2
txtspecial = txtspecial & "Plan X" & vbCrLf
ac = ac + 1
Case Is = 1
txtspecial = txtspecial & "Technique Counter" & vbCrLf & "Trap Sense" & vbCrLf & "Bonus Chakra" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Ninja Police"
ba = ba + b_att(2, Class(19))
will_save = will_save + will(2, Class(19))
ref_save = ref_save + ref(2, Class(19))
fort_save = fort_save + fort(2, Class(19))
hp = hp + (6 + con_mod) * Class(18)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(19)
skill_pts = skill_pts + (6 + int_mod) * Class(19)

Select Case Class(19)
Case Is >= 8
txtspecial = txtspecial & "Sneak Attack + 2d6" & vbCrLf
Case Is >= 4
txtspecial = txtspecial & "Sneak Attack + 1d6" & vbCrLf
End Select

spec_loop = Class(19)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Anticipate" & vbCrLf
Case Is = 9
txtspecial = txtspecial & "Contact - high level" & vbCrLf
ac = ac + 1
Case Is = 8
Case Is = 7
txtspecial = txtspecial & "Bonus Feat - Ninja Police" & vbCrLf
ac = ac + 1
feats = feats + 1
Case Is = 6
txtspecial = txtspecial & "Contact - mid level" & vbCrLf
Case Is = 5
txtspecial = txtspecial & "Bonus Feat - Ninja Police" & vbCrLf
ac = ac + 1
feats = feats + 1
Case Is = 3
txtspecial = txtspecial & "Bonus Feat - Ninja Police" & vbCrLf
ac = ac + 1
feats = feats + 1
Case Is = 2
txtspecial = txtspecial & "Contact - low level" & vbCrLf
Case Is = 1
txtspecial = txtspecial & "Profile" & vbCrLf & "Street savvy" & vbCrLf & "Bonus Chakra" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Ninja Scout"
ba = ba + b_att(2, Class(20))
will_save = will_save + will(2, Class(20))
ref_save = ref_save + ref(1, Class(20))
fort_save = fort_save + fort(3, Class(20))
hp = hp + (8 + con_mod) * Class(20)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(20)
skill_pts = skill_pts + (4 + int_mod) * Class(20)

Select Case Class(20) 'increased speed
Case Is >= 8
txtspecial = txtspecial & "Increased Speed(10 feet)" & vbCrLf
Case Is >= 4
txtspecial = txtspecial & "Increased Speed(5 feet)" & vbCrLf
End Select

Select Case Class(20) 'sneak attack
Case Is = 10
txtspecial = txtspecial & "Sneak Attack +3d6" & vbCrLf
Case Is >= 6
txtspecial = txtspecial & "Sneak Attack +2d6" & vbCrLf
Case Is >= 2
txtspecial = txtspecial & "Sneak Attack +1d6" & vbCrLf
End Select

spec_loop = Class(20)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Quicken technique" & vbCrLf
ac = ac + 1
Case Is = 9
txtspecial = txtspecial & "Bonus Feat - Ninja Scout" & vbCrLf
feats = feats + 1
Case Is = 8
ac = ac + 1
Case Is = 7
txtspecial = txtspecial & "Evasion X" & vbCrLf
ac = ac + 1
Case Is = 6
txtspecial = txtspecial & "Bonus Feat - Ninja Scout" & vbCrLf
feats = feats + 1
Case Is = 5
txtspecial = txtspecial & "Hide in Plain Sight" & vbCrLf
ac = ac + 1
Case Is = 4
ac = ac + 1
Case Is = 3
txtspecial = txtspecial & "Bonus Feat - Ninja Scout" & vbCrLf
feats = feats + 1
Case Is = 2
ac = ac + 1
Case Is = 1
txtspecial = txtspecial & "Track" & vbCrLf & "Bonus Chakra" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Puppeteer"
ba = ba + b_att(3, Class(21))
will_save = will_save + will(2, Class(21))
ref_save = ref_save + ref(2, Class(21))
fort_save = fort_save + fort(2, Class(21))
hp = hp + (6 + con_mod) * Class(18)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(21)
skill_pts = skill_pts + (4 + int_mod) * Class(21)

spec_loop = Class(21)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Puppeteer Skill" & vbCrLf
Case Is = 9
txtspecial = txtspecial & "Advanced Puppetry IV" & vbCrLf
ac = ac + 1
Case Is = 8
txtspecial = txtspecial & "Puppeteer Skill" & vbCrLf
Case Is = 7
txtspecial = txtspecial & "Advanced Puppetry III" & vbCrLf
ac = ac + 1
Case Is = 6
txtspecial = txtspecial & "Puppeteer Skill" & vbCrLf
Case Is = 5
txtspecial = txtspecial & "Advanced Puppetry II" & vbCrLf
ac = ac + 1
Case Is = 4
txtspecial = txtspecial & "Puppeteer Skill" & vbCrLf
Case Is = 3
txtspecial = txtspecial & "Advanced Puppetry" & vbCrLf
ac = ac + 1
Case Is = 2
txtspecial = txtspecial & "Puppeteer Skill" & vbCrLf
Case Is = 1
txtspecial = txtspecial & "Puppetry" & vbCrLf & "Bonus Chakra" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Sacred Fist"
ba = ba + b_att(2, Class(22))
will_save = will_save + will(1, Class(22))
ref_save = ref_save + ref(1, Class(22))
fort_save = fort_save + fort(1, Class(22))
hp = hp + (8 + con_mod) * Class(22)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(22)
skill_pts = skill_pts + (2 + int_mod) * Class(22)

Select Case Class(22)
Case Is >= 9
txtspecial = txtspecial & "Sacred Fist Stance (1d10)" & vbCrLf
Case Is >= 5
txtspecial = txtspecial & "Sacred Fist Stance (1d8)" & vbCrLf
Case Is >= 1
txtspecial = txtspecial & "Sacred Fist Stance (1d6)" & vbCrLf
End Select

spec_loop = Class(22)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Ageless Body" & vbCrLf
Case Is = 9
txtspecial = txtspecial & "Buddhist Palm (Dark iron)" & vbCrLf
ac = ac + 1
Case Is = 8
txtspecial = txtspecial & "Bonus Feat - Sacred Fist" & vbCrLf
feats = feats + 1
Case Is = 7
txtspecial = txtspecial & "Devotion" & vbCrLf
ac = ac + 1
Case Is = 6
txtspecial = txtspecial & "Buddhist Palm (Chakra)" & vbCrLf
Case Is = 5
ac = ac + 1
Case Is = 4
txtspecial = txtspecial & "Evasion" & vbCrLf
Case Is = 3
txtspecial = txtspecial & "Enlightened Defense" & vbCrLf
ac = ac + 1
Case Is = 2
txtspecial = txtspecial & "Bonus Feat - Sacred Fist" & vbCrLf
feats = feats + 1
Case Is = 1
txtspecial = txtspecial & "Bonus Chakra" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Samurai"
ba = ba + b_att(1, Class(23))
will_save = will_save + will(2, Class(23))
ref_save = ref_save + ref(2, Class(23))
fort_save = fort_save + fort(1, Class(23))
hp = hp + (10 + con_mod) * Class(23)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(23)
skill_pts = skill_pts + (2 + int_mod) * Class(23)

spec_loop = Class(23)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Frightful Presence" & vbCrLf
Case Is = 9
txtspecial = txtspecial & "Supreme Path" & vbCrLf & "Bonus Feat - Samurai" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 8
txtspecial = txtspecial & "Weapon Specialisation" & vbCrLf
Case Is = 7
txtspecial = txtspecial & "Armored to the Teeth" & vbCrLf
ac = ac + 1
Case Is = 6
txtspecial = txtspecial & "Bonus Feat - Samurai" & vbCrLf
feats = feats + 1
Case Is = 5
txtspecial = txtspecial & "Sacred Path" & vbCrLf
ac = ac + 1
Case Is = 4
txtspecial = txtspecial & "Combat Acumen" & vbCrLf
Case Is = 3
txtspecial = txtspecial & "Bonus Feat - Samurai" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 2
txtspecial = txtspecial & "Weapon Focus" & vbCrLf
Case Is = 1
txtspecial = txtspecial & "Traditional Path" & vbCrLf & "Bonus Chakra" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Shade"
ba = ba + b_att(2, Class(24))
will_save = will_save + will(3, Class(24))
ref_save = ref_save + ref(1, Class(24))
fort_save = fort_save + fort(3, Class(24))
hp = hp + (6 + con_mod) * Class(24)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(24)
skill_pts = skill_pts + (4 + int_mod) * Class(24)

Select Case Class(24) 'sneak attack
Case Is = 5
txtspecial = txtspecial & "Sneak Attack +4d6" & vbCrLf
Case Is >= 4
txtspecial = txtspecial & "Sneak Attack +3d6" & vbCrLf
Case Is >= 3
txtspecial = txtspecial & "Sneak Attack +2d6" & vbCrLf
Case Is >= 2
txtspecial = txtspecial & "Sneak Attack +1d6" & vbCrLf
End Select

Select Case Class(24) 'save against poison
Case Is = 5
txtspecial = txtspecial & "+3 save against Poison" & vbCrLf
Case Is >= 3
txtspecial = txtspecial & "+2 save against Poison" & vbCrLf
Case Is >= 1
txtspecial = txtspecial & "+1 save against Poison" & vbCrLf
End Select

spec_loop = Class(24)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 5
txtspecial = txtspecial & "Sure Kill" & vbCrLf
ac = ac + 1
Case Is = 2
ac = ac + 1
Case Is = 1
txtspecial = txtspecial & "Death Attack" & vbCrLf & "Poison Expert" & vbCrLf
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Shinobi Adept"
ba = ba + b_att(3, Class(25))
will_save = will_save + will(2, Class(25))
ref_save = ref_save + ref(2, Class(25))
fort_save = fort_save + fort(2, Class(25))
hp = hp + (6 + con_mod) * Class(25)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(25)
skill_pts = skill_pts + (6 + int_mod) * Class(25)

spec_loop = Class(25)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 3
txtspecial = txtspecial & "Chakra Endurance" & vbCrLf & "Evasion" & vbCrLf
ac = ac + 1
Case Is = 2
txtspecial = txtspecial & "Technique Adept" & vbCrLf & "Bonus Feat - Shinobi Adept" & vbCrLf
feats = feats + 1
Case Is = 1
txtspecial = txtspecial & "Combat Tactics" & vbCrLf & "Bonus Chakra" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Shinobi Bodyguard"
ba = ba + b_att(2, Class(26))
will_save = will_save + will(3, Class(26))
ref_save = ref_save + ref(3, Class(26))
fort_save = fort_save + fort(1, Class(26))
hp = hp + (12 + con_mod) * Class(26)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(26)
skill_pts = skill_pts + (2 + int_mod) * Class(26)

Select Case Class(26) 'damage reduction
Case Is >= 9
txtspecial = txtspecial & "Damage Reduction 2/-" & vbCrLf
Case Is >= 5
txtspecial = txtspecial & "Damage Reduction 1/-" & vbCrLf
End Select

spec_loop = Class(26)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Undying Shinobi" & vbCrLf
Case Is = 9
ac = ac + 1
Case Is = 8
txtspecial = txtspecial & "Bonus Feat - Shinobi Bodyguard" & vbCrLf
feats = feats + 1
Case Is = 7
txtspecial = txtspecial & "Mettle" & vbCrLf
ac = ac + 1
Case Is = 6
txtspecial = txtspecial & "Sneak Attack +1d6" & vbCrLf
Case Is = 5
ac = ac + 1
Case Is = 4
txtspecial = txtspecial & "Shinobi's Toughness" & vbCrLf
Case Is = 3
txtspecial = txtspecial & "Bonus Feat - Shinobi Bodyguard" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 2
txtspecial = txtspecial & "Harm's Way" & vbCrLf
Case Is = 1
txtspecial = txtspecial & "Remain Conscious" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Shinobi Swordsman"
ba = ba + b_att(1, Class(27))
will_save = will_save + will(3, Class(27))
ref_save = ref_save + ref(3, Class(27))
fort_save = fort_save + fort(1, Class(27))
hp = hp + (10 + con_mod) * Class(27)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(27)
skill_pts = skill_pts + (2 + int_mod) * Class(27)

Select Case Class(27) 'sneak attack
Case Is >= 8
txtspecial = txtspecial & "Sneak Attack +2d6" & vbCrLf
Case Is >= 4
txtspecial = txtspecial & "Sneak Attack +1d6" & vbCrLf
End Select

spec_loop = Class(27)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Greater Weapon Specialisation" & vbCrLf
Case Is = 9
txtspecial = txtspecial & "Improved Critical" & vbCrLf & "Bonus Feat - Shinobi Swordsman" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 7
txtspecial = txtspecial & "Power of the Elite" & vbCrLf
ac = ac + 1
Case Is = 6
txtspecial = txtspecial & "Greater Weapon Focus" & vbCrLf & "Bonus Feat - Shinobi Swordsman" & vbCrLf
feats = feats + 1
Case Is = 5
txtspecial = txtspecial & "Weapon Specialisation" & vbCrLf
ac = ac + 1
Case Is = 3
txtspecial = txtspecial & "Invisible Strike" & vbCrLf & "Bonus Feat - Shinobi Swordsman" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 2
txtspecial = txtspecial & "Quick Draw" & vbCrLf
Case Is = 1
txtspecial = txtspecial & "Weapon Focus" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Shuriken Expert"
ba = ba + b_att(2, Class(28))
will_save = will_save + will(3, Class(28))
ref_save = ref_save + ref(1, Class(28))
fort_save = fort_save + fort(3, Class(28))
hp = hp + (6 + con_mod) * Class(28)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(28)
skill_pts = skill_pts + (4 + int_mod) * Class(28)

Select Case Class(28) 'sneak attack
Case Is >= 8
txtspecial = txtspecial & "Sneak Attack +2d6" & vbCrLf
Case Is >= 3
txtspecial = txtspecial & "Sneak Attack +1d6" & vbCrLf
End Select

Select Case Class(28) 'precision
Case Is = 10
txtspecial = txtspecial & "Precision +2d4" & vbCrLf
Case Is >= 7
txtspecial = txtspecial & "Precision +1d4" & vbCrLf
End Select

spec_loop = Class(28)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Bullseye" & vbCrLf
ac = ac + 1
Case Is = 9
txtspecial = txtspecial & "Precise Throw" & vbCrLf & "Bonus Feat - Shuriken Expert" & vbCrLf
feats = feats + 1
Case Is = 8
ac = ac + 1
Case Is = 7
txtspecial = txtspecial & "Thrown Weapon Specialisation" & vbCrLf
ac = ac + 1
Case Is = 6
txtspecial = txtspecial & "Greater Thrown Weapon Focus" & vbCrLf & "Bonus Feat - Shuriken Expert" & vbCrLf
feats = feats + 1
Case Is = 5
ac = ac + 1
Case Is = 4
txtspecial = txtspecial & "Quick Draw" & vbCrLf
ac = ac + 1
Case Is = 3
txtspecial = txtspecial & "Bonus Feat - Shuriken Expert" & vbCrLf
feats = feats + 1
Case Is = 2
txtspecial = txtspecial & "Rapid Shot" & vbCrLf
ac = ac + 1
Case Is = 1
txtspecial = txtspecial & "Thrown Weapon Focus" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Soul Edge"
ba = ba + b_att(1, Class(29))
will_save = will_save + will(1, Class(29))
ref_save = ref_save + ref(2, Class(29))
fort_save = fort_save + fort(3, Class(29))
hp = hp + (10 + con_mod) * Class(29)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(29)
skill_pts = skill_pts + (2 + int_mod) * Class(29)

Select Case Class(29) 'increase speed
Case Is >= 8
txtspecial = txtspecial & "Increase Speed (10 feet)" & vbCrLf
Case Is >= 4
txtspecial = txtspecial & "Increase Speed (5 feet)" & vbCrLf
End Select

spec_loop = Class(29)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Ghost Edge" & vbCrLf
Case Is = 9
txtspecial = txtspecial & "Bonus Feat - Soul Edge" & vbCrLf & "Empower Soul Edge (Greater)" & vbCrLf
feats = feats + 1
Case Is = 7
txtspecial = txtspecial & "Shape Soul Edge (Bastard Sword)" & vbCrLf
Case Is = 6
txtspecial = txtspecial & "Bonus Feat - Soul Edge" & vbCrLf & "Empower Soul Edge (Superior)" & vbCrLf
feats = feats + 1
Case Is = 5
txtspecial = txtspecial & "Swift Blade" & vbCrLf & "shape Soul Edge" & vbCrLf
Case Is = 3
txtspecial = txtspecial & "Bonus Feat - Soul Edge" & vbCrLf
feats = feats + 1
Case Is = 2
txtspecial = txtspecial & "Empower Soul Edge (Minor)" & vbCrLf & "Bonus Chakra" & vbCrLf
Case Is = 1
txtspecial = txtspecial & "Soul Edge" & vbCrLf & "Weapon Focus (Soul Edge)" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Squad Captain"
ba = ba + b_att(2, Class(30))
will_save = will_save + will(2, Class(30))
ref_save = ref_save + ref(3, Class(30))
fort_save = fort_save + fort(2, Class(30))
hp = hp + (8 + con_mod) * Class(30)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(30)
skill_pts = skill_pts + (4 + int_mod) * Class(30)

Select Case Class(30) 'sneak attack
Case Is >= 8
txtspecial = txtspecial & "Sneak Attack +3d6" & vbCrLf
Case Is >= 5
txtspecial = txtspecial & "Sneak Attack +2d6" & vbCrLf
Case Is >= 2
txtspecial = txtspecial & "Sneak Attack +1d6" & vbCrLf
End Select

spec_loop = Class(30)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Leading by Example" & vbCrLf
Case Is = 9
txtspecial = txtspecial & "Bonus Feat - Squad Captain" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 8
txtspecial = txtspecial & "Tactical Mastery" & vbCrLf
Case Is = 7
txtspecial = txtspecial & "Improved Command" & vbCrLf & "Mettle" & vbCrLf
ac = ac + 1
Case Is = 6
txtspecial = txtspecial & "Bonus Feat - Squad Captain" & vbCrLf
feats = feats + 1
Case Is = 5
ac = ac + 1
Case Is = 4
txtspecial = txtspecial & "Tactical Expertise" & vbCrLf
Case Is = 3
txtspecial = txtspecial & "Bonus Feat - Squad Captain" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 2
txtspecial = txtspecial & "Force March" & vbCrLf
Case Is = 1
txtspecial = txtspecial & "Command" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Summoner"
ba = ba + b_att(3, Class(31))
will_save = will_save + will(3, Class(31))
ref_save = ref_save + ref(2, Class(31))
fort_save = fort_save + fort(2, Class(31))
hp = hp + (6 + con_mod) * Class(31)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(31)
skill_pts = skill_pts + (2 + int_mod) * Class(31)

Select Case Class(31) 'might of the summoner
Case Is >= 6
txtspecial = txtspecial & "Might of the Summoner +2" & vbCrLf
Case Is >= 2
txtspecial = txtspecial & "Might of the Summoner +1" & vbCrLf
End Select

spec_loop = Class(31)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop

Case Is = 7
txtspecial = txtspecial & "Pride of the Summoner" & vbCrLf & "Bonus Feat - Summoner" & vbCrLf
feats = feats + 1
Case Is = 6
txtspecial = txtspecial & "Will of the Summoner" & vbCrLf
Case Is = 5
txtspecial = txtspecial & "Bonus Feat - Summoner" & vbCrLf
feats = feats + 1
Case Is = 4
txtspecial = txtspecial & "Extend Summoning" & vbCrLf
Case Is = 3
txtspecial = txtspecial & "Bonus Feat - Summoner" & vbCrLf
feats = feats + 1
Case Is = 1
txtspecial = txtspecial & "Empower Summoning" & vbCrLf & "Bonus Chakra" & vbCrLf
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Sword Savant"
ba = ba + b_att(2, Class(32))
will_save = will_save + will(2, Class(32))
ref_save = ref_save + ref(2, Class(32))
fort_save = fort_save + fort(3, Class(32))
hp = hp + (8 + con_mod) * Class(32)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(32)
skill_pts = skill_pts + (2 + int_mod) * Class(32)

spec_loop = Class(32)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop

Case Is = 7
txtspecial = txtspecial & "Sealing Sword (Superior)" & vbCrLf
ac = ac + 1
Case Is = 6
txtspecial = txtspecial & "Heightened Chakra State" & vbCrLf & "Bonus Feat - Sword Savant" & vbCrLf
feats = feats + 1
Case Is = 5
txtspecial = txtspecial & "Weapon Specialisation" & vbCrLf
ac = ac + 1
Case Is = 4
txtspecial = txtspecial & "Sealing Sword (Minor)" & vbCrLf
Case Is = 3
txtspecial = txtspecial & "Sword Weaving" & vbCrLf
ac = ac + 1
Case Is = 2
txtspecial = txtspecial & "Bonus Feat - Sword Savant" & vbCrLf
feats = feats + 1
Case Is = 1
txtspecial = txtspecial & "Chakra State" & vbCrLf & "Weapon Focus" & vbCrLf & "Bonus Chakra" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Taijutsu Master"
ba = ba + b_att(1, Class(33))
will_save = will_save + will(3, Class(33))
ref_save = ref_save + ref(2, Class(33))
fort_save = fort_save + fort(1, Class(33))
hp = hp + (10 + con_mod) * Class(33)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(33)
skill_pts = skill_pts + (2 + int_mod) * Class(33)

spec_loop = Class(33)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 10
txtspecial = txtspecial & "Taijutsu Mastery" & vbCrLf & "Unarmed Attack" & vbCrLf
ac = ac + 1
Case Is = 9
txtspecial = txtspecial & "Bonus Feat - Taijutsu Master" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 8
txtspecial = txtspecial & "Taijutsu Mastery" & vbCrLf
Case Is = 7
txtspecial = txtspecial & "Unarmed Attack" & vbCrLf
ac = ac + 1
Case Is = 6
txtspecial = txtspecial & "Taijutsu Mastery" & vbCrLf
ac = ac + 1
Case Is = 5
txtspecial = txtspecial & "Sneak Attack +1d6" & vbCrLf
ac = ac + 1
Case Is = 4
txtspecial = txtspecial & "Unarmed Attack" & vbCrLf & vbCrLf & "Taijutsu Mastery" & vbCrLf
ac = ac + 1
Case Is = 3
txtspecial = txtspecial & "Bonus Feat - Taijutsu Master" & vbCrLf
feats = feats + 1
Case Is = 2
txtspecial = txtspecial & "Taijutsu Mastery" & vbCrLf
ac = ac + 1
Case Is = 1
txtspecial = txtspecial & "Unarmed Attack" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Technique Analyst"
ba = ba + b_att(3, Class(34))
will_save = will_save + will(1, Class(34))
ref_save = ref_save + ref(3, Class(34))
fort_save = fort_save + fort(3, Class(34))
hp = hp + (6 + con_mod) * Class(34)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(34)
skill_pts = skill_pts + (4 + int_mod) * Class(34)

spec_loop = Class(34)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop

Case Is = 5
txtspecial = txtspecial & "Meta-Chakra Specialisation" & vbCrLf & "Meta-Chakra Application" & vbCrLf
Case Is = 4
txtspecial = txtspecial & "Bonus Feat - Technique Analyst" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 3
txtspecial = txtspecial & "Meta-Chakra Specialisation" & vbCrLf
Case Is = 2
txtspecial = txtspecial & "Bonus Feat - Technique Analyst" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 1
txtspecial = txtspecial & "Meta-Chakra Specialisation" & vbCrLf & "Chakra Theory" & vbCrLf & "Bonus Chakra" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0


Case Is = "Weapon Master"
ba = ba + b_att(1, Class(35))
will_save = will_save + will(3, Class(35))
ref_save = ref_save + ref(3, Class(35))
fort_save = fort_save + fort(2, Class(35))
hp = hp + (10 + con_mod) * Class(35)
clslvls = clslvls & Left(select_cls, 4) & "/" & Class(35)
skill_pts = skill_pts + (2 + int_mod) * Class(35)

spec_loop = Class(35)
Do 'loop will lower by one each time, until it reaches zero, in order to minimise code.

Select Case spec_loop
Case Is = 5
txtspecial = txtspecial & "Improved Critical" & vbCrLf
ac = ac + 1
Case Is = 4
txtspecial = txtspecial & "Greater Weapon Specialisation" & vbCrLf
Case Is = 3
txtspecial = txtspecial & "Greater Weapon Focus" & vbCrLf & "Bonus Feat - Weapon Master" & vbCrLf
feats = feats + 1
ac = ac + 1
Case Is = 2
txtspecial = txtspecial & "Weapon Specialisation" & vbCrLf
Case Is = 1
txtspecial = txtspecial & "Weapon Focus" & vbCrLf
ac = ac + 1
End Select

spec_loop = spec_loop - 1
Loop Until spec_loop = 0

End Select

Next increment

counter(3) = counter(3) + 1

Loop Until select_cls = ""

txtmods = "BA" & vbTab & ba & vbCrLf & "HP" & vbTab & hp & vbCrLf & "WILL" & vbTab & will_save & vbCrLf & "FORT" & vbTab & fort_save & vbCrLf & "REF" & vbTab & ref_save
frmcharsheet.txtmisc = frmcharsheet.txtmisc & "Classes" & vbTab & vbTab & clslvls
End Sub

Private Sub cmdget_info_Click() 'gets the necessary information from the previous form, and then shows character class options.

Dim system As String

root = InputBox("insert the letter where you made your save folder")

Do
If Dir(root & ":", vbDirectory) = "" Then
root = InputBox("Error, please enter a valid root directory.")
Else
End If
Loop Until Dir(root & ":", vbDirectory) <> ""

 Filename = root & ":\character_creator\system"
 Open Filename For Input As #1
 system = Input$(LOF(1), 1)
Close #1

Select Case system 'vbcrlf is used due to it being added to the variable after being saved. unsure why.
Case Is = "D&D" & vbCrLf
txtdisplay = "You are using the D&D 3.5 rules and system. A list of available classes are opposite. Please note that it is impossible to prestige before level 4, so a character must have at least 3 levels in a basic class in order to meet any of the prerequisites. The prereq's for each class (if any) are indented under their names."
Call dnd_info

Case Is = "D20 Modern" & vbCrLf
txtdisplay = "You are using the D20 Modern rules and system. A list of available classes are opposite. Please note that it is impossible to prestige before level 4, so a character must have at least 3 levels in a basic class in order to meet any of the prerequisites. The prereq's for each class (if any) are indented under their names."
Call mod_info

Case Is = "Naruto D20" & vbCrLf
txtdisplay = "You are using the Naruto D20 rules and system. A list of available classes are opposite. Please note that it is impossible to prestige before level 4, so a character must have at least 3 levels in a basic class in order to meet any of the prerequisites. The prereq's for each class (if any) are indented under their names."
Call naruto_info
End Select

End Sub

Sub dnd_info() 'shows the class choices in the dnd system

txtclasses = "Basic" & vbCrLf & vbCrLf & "Barbarian" & vbCrLf & "Bard" & vbCrLf & "Cleric" & vbCrLf & "Druid" & vbCrLf & "Fighter" & vbCrLf & "Monk" & vbCrLf & "Paladin" & vbCrLf & "Ranger" & vbCrLf & "Rogue" & vbCrLf & "Sorcerer" & vbCrLf & "Wizard"
txtprestclasses = txtprestclasses & "Prestige" & vbCrLf & vbCrLf & "Arcane Archer:" & vbCrLf & "  -  Elf or Half-elf, BA+6, Point Blank Shot, Precise Shot, Weapon Focus(long or shortbow), able to cast level 1 arcane spells" & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Arcane Trickster:" & vbCrLf & "  -  Any nonlawful, Decipher Script 7, Disable Device 7, Escape Artist 7, Knowledge Arcana 4, able to cast mage hand and at least one 3+ level arcane spell, sneak attack +2d6" & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Archmage:" & vbCrLf & "  -  Knowledge Arcana 15, spellcraft 15, Skill Focus(Spellcraft), Spell Focus in two schools of magic, able to cast level 7 arcane spells, and know level 5+ spells from at least 5 schools of magic" & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Assassin:" & vbCrLf & "  -  Any Evil, Disguise 4, Hide 8, Move Silently 8, Kill someone for no other reason than to join the assassins. " & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Blackguard:" & vbCrLf & "  -  Any Evil, BA +6, Hide 5, Knowledge(Religion) 2, cleave, Improved Sunder, Power Attack, Made peaceful contact with an Evil Outsider who was summoned by him or someone else" & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Dragon Disciple:" & vbCrLf & "  -  Any non-dragon, Knowledge(Arcana) 8, Draconic, can cast arcane spells without preperation, choose a variety at first level in this class" & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Duelist:" & vbCrLf & "  -  BA +6, perform 3, tumble 5, Dodge, Mobility, Weapon Finesse." & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Dwarven Defender:" & vbCrLf & "  -  Dwarf, Any Lawful, BA+7, Dodge, Endurance, Toughness" & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Eldritch Knight:" & vbCrLf & "  -  Proficient with all martial weapons, able to cast level 3 arcane spells" & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Hierophant:" & vbCrLf & "  -  Knowledge(religion) 15, Any metamagic feat, able to cast level 7 divine spells" & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Horizon Walker:" & vbCrLf & "  -  Knowledge(geography) 8, Endurance" & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Loremaster:" & vbCrLf & "  -  Knowledge(Any Two) 10, Any three metamagic or item creation feats, Skill Focus(Knowledge(any)), Able to cast 7 different divination spells, one at least level 3" & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Mystic Theurg:e" & vbCrLf & "  -  Knowledge(Arcana) 6, Knowledge(Religion) 6, able to cast level 2 arcane and divine spells" & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Shadowdancer:" & vbCrLf & "  -  Move Silently 8, Hide 10, perform(Dance) 5, Combat Reflexes, Dodge, Mobility" & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Thaumaturgist:" & vbCrLf & "  -  Spell Focus(Conjuration), able to cast lesser planar ally"
End Sub

Sub mod_info() 'shows the class choices in the modern system

txtclasses = "Basic" & vbCrLf & vbCrLf & "Strong" & vbCrLf & "Fast" & vbCrLf & "Tough" & vbCrLf & "Smart" & vbCrLf & "Dedicated" & vbCrLf & "Charismatic"
txtprestclasses = txtprestclasses & "Prestige" & vbCrLf & vbCrLf & "Soldier:" & vbCrLf & " - BA+3, Knowledge(Tactics) 3, Personal Firearms Proficiency" & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Martial Artist:" & vbCrLf & " - BA+3, Jump 3, Combat Martial Arts, Defensive Martial Arts" & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Gunslinger:" & vbCrLf & " - BA+2, Sleight of Hand 6, Tumble 6, Personal Firearms Proficiency" & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Infiltrator:" & vbCrLf & " - BA+2, Hide 6, Move Silently 6" & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Daredevil:" & vbCrLf & " - BA+2, Concentration 6, Drive 6, Endurance" & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Bodyguard:" & vbCrLf & " - BA+2, Concentrate 6, Intimidate 6, Personal Firearms Proficiency" & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Field Scientist:" & vbCrLf & " - Research 6, Craft(Chemical, or Electronic) 6, Knowledge(Earth and Life Sciences, Physical Sciences, or Technology) 6" & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Techie:" & vbCrLf & " - Computer use 6, Disable Device 6, Craft(Electronic, or Mechanical) 6" & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Field Medic:" & vbCrLf & " - BA+2, Treat Injury 6, Spot 6, Surgery" & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Investigator:" & vbCrLf & " - BA+2, Investigate 6, Listen 6, Sense Motive 6" & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Personality:" & vbCrLf & " - Diplomacy 6, Perform(Any) 6, Renown" & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Negotiator:" & vbCrLf & " - Bluff 6, Diplomacy 6, Alertness" & vbCrLf
End Sub

Sub naruto_info() 'shows the class choices in the naruto system

txtclasses = "Basic" & vbCrLf & vbCrLf & "Strong" & vbCrLf & "Fast" & vbCrLf & "Tough" & vbCrLf & "Smart" & vbCrLf & "Dedicated" & vbCrLf & "Charismatic"
txtprestclasses = txtprestclasses & "Prestige" & vbCrLf & vbCrLf & "Beast Lord:" & vbCrLf & "BA+6, Handle animal 9, Survival 9, Animal Affinity, Moujuu Aishou, Frenzy 1/day or better" & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Beast Master:" & vbCrLf & "BA+2, Handle animal 6, Survival 3, Moujuu Aishou" & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Blink Strike:" & vbCrLf & "Move silently 9, Ninjutsu 12, Taijutsu 12, Tumble 9, Dodge, Mobility, Quick Draw, 5th step of mastery in Shunshin no jutsu, know at least one spacetime w/ teleporter description, and 5 taijutsu tech's." & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Devastator:" & vbCrLf & "Chakra control 16, Know(Nin lore) 12, Genjutsu 12 or Ninjutsu 12, Genjutsu Adept or Ninjutsu Adept, Technique Focus(Any), any two meta-Chakra feats, effective skill threshold of 18 in either ninjutsu or genjutsu, 20 or more bonus reserve from talents, feats, or class levels." & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Elementalist:" & vbCrLf & "Ninjutsu 9, Ninjutsu Adept, know at least 4 ninjutsu tech's of the element he wants to specialize in." & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Exarch:" & vbCrLf & "Chakra control 16, Know(Earth and life science) 12, Ninjutsu 16, Treat injury 12, Harmony, Any meta-chakra feat, Chakra Scalpel and Chakra Scalpel Expertise abilities." & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Exemplar:" & vbCrLf & "Taijutsu 16, Defensive Martial Arts, Taijutsu Adept, Weapon Focus (any weapon) and Archaic Weapons Prof or Combat Martial Arts or Exotic Melee Prof or Ranged Weapon Prof or Nin Weapons Prof, Weapon Specialization class feature, effective skill threshold of 18 in strike taijutsu tech's." & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Genjutsu Master:" & vbCrLf & "Genjutsu 9, Genjutsu Adept, know at least 6 Genjutsu tech's." & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Livewire:" & vbCrLf & "Sleight of hand 6, Weapon Focus(Battle Wire), 3rd step of mastery in Kousen Ryu." & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Master Strategist:" & vbCrLf & "Know(Tactics) 6, Combat Expertise, Plan, or Plan X class feature." & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Medical Specialist:" & vbCrLf & "BA+2, Chakra control 6, Know(Earth and life science) 6, Ninjutsu 6, Treat injury 6, Harmony, Medical Expert." & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Ninja Operations Counter:" & vbCrLf & "Genjutsu 6, Hide 9, Know(Nin lore) 6, Ninjutsu 7, Move silently 9, Survival 6, Stealthy, Track, 3rd step of mastery in 3 rank 4 tech's." & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Ninja Police:" & vbCrLf & "Gather information 3, Investigate 6, Sense motive 3, 8 ranks amongst Chakra control, Genjutsu, Ninjutsu, and Taijutsu, Attentive." & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Ninja Scout:" & vbCrLf & "Know(Nin lore) 6, Survival 3, 10 ranks amongst Chakra control, Genjutsu, Ninjutsu, and Taijutsu, Nin Weapons Prof" & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Puppeteer:" & vbCrLf & "BA+1, Concentration 3, Ninjutsu 6, repair 6, able to create C-class chakra threads with Ninpou:Chakra No Ito." & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Sacred Fist:" & vbCrLf & "BA+2, Chakra control 6, Know(Theology and philosophy) 6, Taijutsu 6, Combat Martial Arts, Harmony." & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Samurai:" & vbCrLf & "BA+3, KNow(Tactics) 6, Taijutsu 6, Armour prof(Light), Archaic Weapons Prof or Exotic Weapons Prof or Nin Weapons Prof." & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Shade:" & vbCrLf & "BA+4, Hide 9, Listen 6, Move silently 9, Spot 6, Alertness or Stealthy, Improved Initiative, able to sense or suppress chakra" & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Shinobi Adept:" & vbCrLf & "Know(Nin lore) 6, Genjutsu or Ninjutsu 6, 3rd step of mastery in one Ninjutsu or genjutsu tech." & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Shinobi Bodyguard:" & vbCrLf & "BA+2, Concentration 6, Great Fortitude." & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Shinobi Swordsman:" & vbCrLf & "BA+3, hide 3, Move silently 3, Taijutsu 3, Stealthy, proficient in chosen weapon, know at least 3 strike or stance taijutsu tech's" & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Shuriken Expert:" & vbCrLf & "BA+3, Sleight of hand 6, Tumble 6, Archaic Weapons Prof or Nin Weapons Prof, Point Blank Shot." & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Soul Edge:" & vbCrLf & "BA+2, Chakra Control 6, Archaic Weapons Prof, 3rd step of mastery in the Seireiha tech." & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Squad Captain:" & vbCrLf & "BA+2, Know(tactics) 6, Diplomacy or Profession 6, Genin, Nin Weapons Prof." & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Summoner:" & vbCrLf & "Base Will+4, Ninjutsu 9, Blood Pact, Retrieval Expert, Kuchiyose No Jutsu tech." & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Sword Savant:" & vbCrLf & "BA+6, Ninjutsu 12, Taijutsu 12, Ninjutsu Adept, Taijutsu Adept, Proficient in any archaic, exotic or nin weapon." & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Taijutsu Master:" & vbCrLf & "BA+3, Taijutsu 6, Combat Martial Arts, know at least 4 taijutsu tech's." & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Technique Analyst:" & vbCrLf & "Chakra control 9, 6 ranks amongst Concentration, Genjutsu, Know(Nin lore) Ninjutsu and Taijutsu." & vbCrLf
txtprestclasses = txtprestclasses & vbCrLf & "Weapon Master:" & vbCrLf & "BA+6, Know(Tactics) 9, Jump or Tumble 9, Taijutsu 9, Archaic Weapons Prof or Combat Martial Arts or Nin Weapons Prof or Exotic melee Weapon Prof, Know at least 6 strike or stance taijutsu tech's." & vbCrLf
End Sub

