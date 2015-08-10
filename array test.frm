VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8850
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15180
   LinkTopic       =   "Form1"
   ScaleHeight     =   8850
   ScaleWidth      =   15180
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtr3 
      Height          =   4335
      Left            =   9240
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtw3 
      Height          =   4335
      Left            =   6720
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtf3 
      Height          =   4335
      Left            =   4200
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtba3 
      Height          =   4335
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtr2 
      Height          =   4335
      Left            =   8520
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtw2 
      Height          =   4335
      Left            =   6000
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtf2 
      Height          =   4335
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtba2 
      Height          =   4335
      Left            =   960
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtr1 
      Height          =   4335
      Left            =   7800
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtw1 
      Height          =   4335
      Left            =   5280
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtf1 
      Height          =   4335
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtba1 
      Height          =   4335
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton cmdend 
      Caption         =   "end"
      Height          =   855
      Left            =   10680
      TabIndex        =   1
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "get arrays"
      Height          =   495
      Left            =   10680
      TabIndex        =   0
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "ref"
      Height          =   255
      Left            =   7920
      TabIndex        =   17
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "will"
      Height          =   375
      Left            =   5400
      TabIndex        =   16
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "fort"
      Height          =   255
      Left            =   2880
      TabIndex        =   15
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "ba"
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdend_Click()
End
End Sub

Private Sub Command1_Click()
'populates save arrays
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


Dim counter As Integer

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

Dim count As Integer

For count = 1 To 20

txtba1 = txtba1 & b_att(1, count) & vbCrLf
txtba2 = txtba2 & b_att(2, count) & vbCrLf
txtba3 = txtba3 & b_att(3, count) & vbCrLf

txtf1 = txtf1 & fort(1, count) & vbCrLf
txtf2 = txtf2 & fort(2, count) & vbCrLf
txtf3 = txtf3 & fort(3, count) & vbCrLf

txtw1 = txtw1 & will(1, count) & vbCrLf
txtw2 = txtw2 & will(2, count) & vbCrLf
txtw3 = txtw3 & will(3, count) & vbCrLf

txtr1 = txtr1 & ref(1, count) & vbCrLf
txtr2 = txtr2 & ref(2, count) & vbCrLf
txtr3 = txtr3 & ref(3, count) & vbCrLf
Next count










End Sub
