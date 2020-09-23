VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4710
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3795
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   3360
      Width           =   4455
   End
   Begin VB.CommandButton cmdUse 
      Caption         =   "Select"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   2640
      Width           =   2535
   End
   Begin VB.FileListBox File1 
      Height          =   2040
      Left            =   2160
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdUse_Click()
If Text2.Text = "Text1" Then
    Form1.Text1.Text = Form2.Text1.Text
ElseIf Text2.Text = "Text2" Then
    Form1.Text2.Text = Form2.Text1.Text
End If
    Form2.Hide
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub File1_Click()
If File1.Path = "C:\" Then
    Text1.Text = File1.Path & File1.FileName
Else
    Text1.Text = File1.Path & "\" & File1.FileName
End If
End Sub

Private Sub Form_Load()
Dir1.Path = "C:\"
End Sub
