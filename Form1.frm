VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "Domain Name Finder"
   ClientHeight    =   6600
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6840
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Text            =   "Available_Domains.txt"
      Top             =   3720
      Width           =   2775
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   4575
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Width           =   4575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Browse"
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   5535
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   4920
      TabIndex        =   3
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Program"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   3840
      Width           =   1815
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   4200
      Width           =   4695
      ExtentX         =   8281
      ExtentY         =   4048
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label Label1 
      Caption         =   "Name of File to Create?"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Choose the path of the output file:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Location of Input Text File of the Domain Names to check:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   5175
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Index           =   0
      Begin VB.Menu Start 
         Caption         =   "Start"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu About 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim doc As IHTMLDocument2
    Dim ElementCol As IHTMLElementCollection
    Dim Element As IHTMLElement
    Dim PageState As Integer
    Dim Domain As String
    Dim objFSO As Object
    Dim objTextFile As Object
    Dim objFSO2 As Object
    Dim objTextFile2 As Object
    Dim FileOpened As Boolean
    Dim FileOpened2 As Boolean
    Dim strLine As String


    
Private Sub Enter_Domain_Name(ByVal Domain)
    For i = 0 To (ElementCol.length - 1)
        Set Element = ElementCol.Item(i)
        If (Element.tagName = "INPUT") Then
            If InStr(Element.outerHTML, "name=vhost_name") > 0 Then
               Element.innerText = Domain
            End If
        End If
    Next i
End Sub

Private Sub Click_Submit()
    For i = 0 To (ElementCol.length - 1)
        Set Element = ElementCol.Item(i)
        If (Element.tagName = "INPUT") Then
            If InStr(UCase(Element.outerHTML), UCase("name=chk_domain")) > 0 Then
               'MsgBox Element.outerHTML
            Element.Click
            End If
        End If
    Next i
End Sub
Private Sub IsDomainAvailable()
Dim Found
Found = False
    For i = 0 To (ElementCol.length - 1)
        Set Element = ElementCol.Item(i)
        If (Element.tagName = "B") Then
            If InStr(UCase(Element.outerHTML), UCase("Taken")) > 0 Then
               'MsgBox "Domain Name is Available"
               Found = True
            End If
        End If
    Next i
If Found Then
    'MsgBox "Domain Name already in use"
    PageState = 2
    'WebBrowser1.GoBack
    Call cmdStart_Click
Else
    PageState = 2
    'Store the Domain name
    List1.AddItem (Domain & ".Com")
    Call StoreDomainName(Domain & ".com")
    'MsgBox "store the domain name"
    'WebBrowser1.GoBack
    Call cmdStart_Click
End If
End Sub
Public Sub StoreDomainName(ByVal DomainName)
    Dim FileLoc As String
If Not FileOpened2 Then
    Const forReading = 1, forWriting = 2, forAppending = 3
    FileLoc = Text2.Text & Text3.Text
    'MsgBox FileLoc
    FileLoc = Replace(FileLoc, "\\", "\")
    
    Set objFSO2 = CreateObject("scripting.FileSystemObject")
    'Open the file
    Set objTextFile2 = objFSO2.CreateTextFile(FileLoc)
    FileOpened2 = True
    objTextFile2.WriteLine "List of available domain names.."
End If
    objTextFile2.WriteLine DomainName
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub About_Click()
Form3.Show
End Sub

Private Sub cmdStart_Click()
If Text1.Text = "" Then
    MsgBox "The Input Text File CANNOT be BLANK"
Else
    If FileOpened Then
        Domain = GetDomainName()
        If ValidChar(Domain) Or PageState = 3 Then
            'MsgBox Domain
            If PageState <> 3 Then
                Call Enter_Domain_Name(Domain)
                PageState = 1
                Call Click_Submit
            Else
                Call CloseObjects
                MsgBox "End of File has been reached!"
            End If
        Else
            Call cmdStart_Click
        End If
    Else
        Call OpenTxtFile
        Call cmdStart_Click
    End If
End If
End Sub
Public Function GetDomainName()
Dim strWord As String
strWord = ""
If Not objTextFile.AtEndOfStream Then
    If strLine = "" Then
        strLine = Trim(objTextFile.Readline)
    End If
    'Parse the line into words
    strWord = getWord(strLine)
    GetDomainName = strWord
Else
    PageState = 3
End If
End Function
Public Function getWord(strLine)
Dim length, pos1, pos2 As Integer
Dim strLeft, strRight As String
strLeft = ""
strRight = ""
length = 0

length = Len(strLine)
'get the the first 2 spaces.
pos1 = InStr(strLine, " ")

If pos1 > 0 Then
    strLeft = Left(strLine, pos1 - 1)
    strRight = Right(strLine, length - pos1)
    'MsgBox "Left: " & strLeft & "| Right: " & strRight & "|"
    strLine = strRight
    getWord = strLeft
Else 'We only have 1 word remaining
    'MsgBox "Line: " & strLine & "| Right: " & strRight & "|"
    getWord = strLine
    strLine = ""
End If
'MsgBox getWord
End Function
Public Function ValidChar(ByVal strWord) As Boolean
Dim ValChar As String
Dim myChar As String
Dim length As Integer

length = Len(strWord)
strWord = LCase(strWord)
ValChar = "1234567890abcdefghijklmnopqrstuvwxyz_-"
ValidChar = True

If length > 2 Then
    For i = 1 To length
        myChar = Mid(strWord, i, 1)
        If InStr(ValChar, myChar) = 0 Then
            ValidChar = False
        End If
    Next
Else
    ValidChar = False
End If

End Function
Public Sub CloseTxtFile()
    objTextFile.Close
    Set objTextFile = Nothing
    Set objFSO = Nothing
End Sub

Public Sub OpenTxtFile()
    Dim FileLoc As String
    Const forReading = 1, forWriting = 2, forAppending = 3
    FileLoc = Text1.Text
    FileLoc = Replace(FileLoc, "\\", "\")
    
    Set objFSO = CreateObject("scripting.FileSystemObject")
    'Open the file
    Set objTextFile = objFSO.OpenTextFile(FileLoc)
    FileOpened = True
End Sub


Private Sub cmdTest2_Click()
If Not FileOpened Then
Call OpenTxtFile
End If

Domain = GetDomainName()

If ValidChar(Domain) Then
    MsgBox Domain
Else
    Call cmdTest2_Click
End If
    
End Sub


Private Sub Command2_Click()
Form2.Show
Form2.Text2.Text = "Text1"
End Sub

Private Sub Command3_Click()
Form2.Show
Form2.Text2.Text = "Text2"
End Sub

Private Sub Dir1_Change()
Text2.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo ER_Drive1_Change
Dir1.Path = Drive1.Drive
Exit Sub

ER_Drive1_Change:
    MsgBox "There ain't no disk in " & Drive1.Drive & "\ Dumb Ass!"
End Sub

Private Sub Form_Load()
Dir1.Path = "C:\"
FileOpened = False
FileOpened2 = False
PageState = 0
cmdStart.Caption = "Please Wait..."
cmdStart.Enabled = False
    WebBrowser1.Navigate "http://www.register.com"
List1.AddItem ("Available Domains: ")
End Sub

Private Sub Label5_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Start_Click()
Call cmdStart_Click
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    Set doc = WebBrowser1.Document
    Set ElementCol = doc.body.All
    Select Case PageState
        Case 0
            cmdStart.Caption = "Start Program"
            cmdStart.Enabled = True
        Case 1
            'MsgBox ("Page Has Loaded")
            Call IsDomainAvailable
        Case 2
            Call cmdStart_Click
    End Select
End Sub
Public Sub CloseObjects()
objTextFile.Close
Set objFSO = Nothing
If FileOpened2 Then
    objTextFile2.Close
    Set objFSO2 = Nothing
End If
End Sub
