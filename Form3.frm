VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5565
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   5820
   ScaleWidth      =   5565
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   5400
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   5295
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form3.frx":0442
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form3.Hide
End Sub

Private Sub Form_Load()
Text1.Text = "Program design specs: "
Text1.Text = Text1.Text & " This program will read input from a text file"
Text1.Text = Text1.Text & " It will then read in line by line."
Text1.Text = Text1.Text & " For each line it will take words which are seperated by spaces."
Text1.Text = Text1.Text & " It will validate each words to make sure that the characters are "
Text1.Text = Text1.Text & "only letters, numbers, and _- characters."
Text1.Text = Text1.Text & " Words with invalid characters will be ignored."
Text1.Text = Text1.Text & " Valid domain names will be stored in a text file specified by the user. Created By Titon Hoque"

End Sub

