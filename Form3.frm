VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select batch file to open"
   ClientHeight    =   5190
   ClientLeft      =   6120
   ClientTop       =   4125
   ClientWidth     =   4230
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   4230
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Open File"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   4680
      Width           =   2295
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      Pattern         =   "*.bat"
      TabIndex        =   2
      Top             =   3120
      Width           =   3975
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2565
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3975
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    OpenFile
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub File1_DblClick()
    OpenFile
End Sub

Private Sub Form_Load()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub OpenFile()
    'the if statement below is like this because i was having problems opening file in just
    'my 'C' Drive (no sub-folders), change this to your default drive
    If File1.Path = "c:\" Then
        Open File1.Path & File1.FileName For Input As #1
            a = Input(LOF(1), 1)
        Close #1
        Form1.Text1.Text = a
        Form3.Visible = False
    Else
        Open File1.Path & "\" & File1.FileName For Input As #1
            a = Input(LOF(1), 1)
        Close #1
        Form1.Text1.Text = a
        Form3.Visible = False
    End If
End Sub
