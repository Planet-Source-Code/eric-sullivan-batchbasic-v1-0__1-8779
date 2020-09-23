VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "[BatchBasic] [v1.0] [By: Eric Sullivan]"
   ClientHeight    =   6990
   ClientLeft      =   3555
   ClientTop       =   3285
   ClientWidth     =   10215
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   10215
   Begin VB.TextBox Text1 
      Height          =   6735
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   9975
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
      Begin VB.Menu MnuNB 
         Caption         =   "&New Batch "
         Shortcut        =   ^N
      End
      Begin VB.Menu MnuOB 
         Caption         =   "&Open Batch"
         Shortcut        =   ^O
      End
      Begin VB.Menu MnuBRK1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuP 
         Caption         =   "&Print "
         Shortcut        =   ^P
      End
      Begin VB.Menu MnuBRK3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRB 
         Caption         =   "&Run Batch"
         Shortcut        =   ^R
      End
      Begin VB.Menu MnuCB 
         Caption         =   "&Compile Batch"
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuBRK4 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu MnuView 
      Caption         =   "&View"
      Begin VB.Menu MnuScomm 
         Caption         =   "Small Commands"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuLcomm 
         Caption         =   "Large Commands"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu MnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu SMnuHelp 
         Caption         =   "Help"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Form2.Visible = True
    Form5.Visible = True
    Form5.List1.Selected(0) = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    QuitMessage
End Sub

Private Sub MnuCB_Click()
    Open App.Path & "\Compiled\BatchBasicRunBatch.bat" For Output As #1
        Print #1, Text1.Text
    Close #1
    MsgBox "Batch File Has Been Compiled And Placed In The Subfolder Called Compiled", vbOKOnly, "BatchBasic v1.0"
End Sub

Private Sub MnuExit_Click()
    QuitMessage
End Sub

Private Sub MnuLcomm_Click()
    If MnuLcomm.Checked = True Then
        Form5.Visible = False
        MnuLcomm.Checked = False
    ElseIf MnuLcomm.Checked = False Then
        Form5.Visible = True
        MnuLcomm.Checked = True
    End If
End Sub

Private Sub MnuNB_Click()
    Text1.Text = ""
End Sub

Private Sub MnuOB_Click()
    Form3.Visible = True
End Sub

Private Sub MnuP_Click()
    Printer.NewPage
    Printer.Print "BatchBasic Print-off" & vbNewLine & vbNewLine & Text1.Text
    Printer.EndDoc
End Sub

Private Sub MnuRB_Click()
    Open App.Path & "\Compiled\BatchBasicRunBatch.bat" For Output As #1
        Print #1, Text1.Text
    Close #1
    Shell App.Path & "\Compiled\BatchBasicRunBatch.bat", vbNormalFocus
End Sub

Private Sub QuitMessage()
    QExit = MsgBox( _
       "Are you sure you would like to quit?", _
       vbYesNo + vbQuestion, _
       "BatchBasic v1.0")
    
    Select Case QExit
       Case vbYes
          End
       Case vbNo
          Cancel = Not ReadyToQuit
    End Select
End Sub

Private Sub MnuScomm_Click()
    If MnuScomm.Checked = True Then
        Form2.Visible = False
        MnuScomm.Checked = False
    ElseIf MnuScomm.Checked = False Then
        Form2.Visible = True
        MnuScomm.Checked = True
    End If
End Sub

Private Sub SMnuHelp_Click()
    Form7.Visible = True
End Sub
