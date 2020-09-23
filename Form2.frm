VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "         Add Commands"
   ClientHeight    =   7260
   ClientLeft      =   945
   ClientTop       =   3000
   ClientWidth     =   2520
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   2520
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   6360
      Width           =   2295
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5910
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add Commands"
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
      Left            =   120
      TabIndex        =   0
      Top             =   6840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   6120
      Width           =   2295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If List1.Selected(0) = True Then
        If Form1.Text1.Text = "" Then
            Form1.Text1.SelText = "@echo on"
        ElseIf Form1.Text1.Text <> "" Then
            Form1.Text1.SelText = vbNewLine & "@echo on"
        End If
    ElseIf List1.Selected(1) = True Then
        If Form1.Text1.Text = "" Then
            Form1.Text1.SelText = "@echo off"
        ElseIf Form1.Text1.Text <> "" Then
            Form1.Text1.SelText = vbNewLine & "@echo off"
        End If
    ElseIf List1.Selected(3) = True Then
        If Form1.Text1.Text = "" Then
            Form1.Text1.SelText = "exit"
        ElseIf Form1.Text1.Text <> "" Then
            Form1.Text1.SelText = vbNewLine & "exit"
        End If
    ElseIf List1.Selected(4) = True Then
        If Form1.Text1.Text = "" Then
            Form1.Text1.SelText = "C:\WINDOWS\RUNDLL.EXE user.exe,exitwindowsexec"
        ElseIf Form1.Text1.Text <> "" Then
            Form1.Text1.SelText = vbNewLine & "C:\WINDOWS\RUNDLL.EXE user.exe,exitwindowsexec"
        End If
    ElseIf List1.Selected(6) = True Then
        If Form1.Text1.Text = "" Then
            Form1.Text1.SelText = "START /M " & Chr(34) & Text1.Text & Chr(34)
        ElseIf Form1.Text1.Text <> "" Then
            Form1.Text1.SelText = vbNewLine & "START /M " & Chr(34) & Text1.Text & Chr(34)
        End If
    ElseIf List1.Selected(8) = True Then
        If Form1.Text1.Text = "" Then
            Form1.Text1.SelText = "rem " & Text1.Text
        ElseIf Form1.Text1.Text <> "" Then
            Form1.Text1.SelText = vbNewLine & "rem " & Text1.Text
        End If
    ElseIf List1.Selected(9) = True Then
        If Form1.Text1.Text = "" Then
            Form1.Text1.SelText = "echo " & Text1.Text
        ElseIf Form1.Text1.Text <> "" Then
            Form1.Text1.SelText = vbNewLine & "echo " & Text1.Text
        End If
    ElseIf List1.Selected(10) = True Then
        If Form1.Text1.Text = "" Then
            Form1.Text1.SelText = "echo."
        ElseIf Form1.Text1.Text <> "" Then
            Form1.Text1.SelText = vbNewLine & "echo."
        End If
    ElseIf List1.Selected(11) = True Then
        If Form1.Text1.Text = "" Then
            Form1.Text1.SelText = "cls"
        ElseIf Form1.Text1.Text <> "" Then
            Form1.Text1.SelText = vbNewLine & "cls"
        End If
    ElseIf List1.Selected(13) = True Then
        If Form1.Text1.Text = "" Then
            Form1.Text1.SelText = "pause"
        ElseIf Form1.Text1.Text <> "" Then
            Form1.Text1.SelText = vbNewLine & "pause"
        End If
    ElseIf List1.Selected(14) = True Then
        If Form1.Text1.Text = "" Then
            Form1.Text1.SelText = "echo " & Text1.Text & vbNewLine & "pause > nul"
        ElseIf Form1.Text1.Text <> "" Then
            Form1.Text1.SelText = vbNewLine & "echo " & Text1.Text & vbNewLine & "pause > nul"
        End If
    ElseIf List1.Selected(15) = True Then
        If Form1.Text1.Text = "" Then
            Form1.Text1.SelText = "type nul | choice.com /n /cy /ty," & Text1.Text & "> nul"
        ElseIf Form1.Text1.Text <> "" Then
            Form1.Text1.SelText = vbNewLine & "type nul | choice.com /n /cy /ty," & Text1.Text & "> nul"
        End If
    ElseIf List1.Selected(17) = True Then
        If Form1.Text1.Text = "" Then
            Form1.Text1.SelText = "del " & Text1.Text
        ElseIf Form1.Text1.Text <> "" Then
            Form1.Text1.SelText = vbNewLine & "del " & Text1.Text
        End If
    ElseIf List1.Selected(18) = True Then
        If Form1.Text1.Text = "" Then
            Form1.Text1.SelText = "echo y | del " & Text1.Text & " > nul"
        ElseIf Form1.Text1.Text <> "" Then
            Form1.Text1.SelText = vbNewLine & "echo y | del " & Text1.Text & " > nul"
        End If
    ElseIf List1.Selected(20) = True Then
        If Form1.Text1.Text = "" Then
            Form1.Text1.SelText = "md " & Text1.Text
        ElseIf Form1.Text1.Text <> "" Then
            Form1.Text1.SelText = vbNewLine & "md " & Text1.Text
        End If
    ElseIf List1.Selected(21) = True Then
        If Form1.Text1.Text = "" Then
            Form1.Text1.SelText = "rd " & Text1.Text
        ElseIf Form1.Text1.Text <> "" Then
            Form1.Text1.SelText = vbNewLine & "rd " & Text1.Text
        End If
    ElseIf List1.Selected(23) = True Then
        If Form1.Text1.Text = "" Then
            Form1.Text1.SelText = "prompt " & Text1.Text
        ElseIf Form1.Text1.Text <> "" Then
            Form1.Text1.SelText = vbNewLine & "prompt " & Text1.Text
        End If
    ElseIf List1.Selected(25) = True Then
        If Form1.Text1.Text = "" Then
            Form1.Text1.SelText = "date"
        ElseIf Form1.Text1.Text <> "" Then
            Form1.Text1.SelText = vbNewLine & "date"
        End If
    ElseIf List1.Selected(26) = True Then
        If Form1.Text1.Text = "" Then
            Form1.Text1.SelText = "time"
        ElseIf Form1.Text1.Text <> "" Then
            Form1.Text1.SelText = vbNewLine & "time"
        End If
    End If
    Text1.Text = ""
End Sub

Private Sub Form_Load()
    Label1.Visible = False
    Text1.Visible = False
    
    List1.AddItem ("@echo on")
    List1.AddItem ("@echo off")
    List1.AddItem ("------------------------------")
    List1.AddItem ("Exit Batch File")
    List1.AddItem ("Exit Windows")
    List1.AddItem ("------------------------------")
    List1.AddItem ("Run Program")
    List1.AddItem ("------------------------------")
    List1.AddItem ("Add Rem Statement")
    List1.AddItem ("Add Text")
    List1.AddItem ("insert space")
    List1.AddItem ("Clear Screen")
    List1.AddItem ("------------------------------")
    List1.AddItem ("Original Pause")
    List1.AddItem ("Custom Pause")
    List1.AddItem ("Add Delay")
    List1.AddItem ("------------------------------")
    List1.AddItem ("Delete file")
    List1.AddItem ("Delete file (no prompt)")
    List1.AddItem ("------------------------------")
    List1.AddItem ("Make Directory")
    List1.AddItem ("Remove Directory")
    List1.AddItem ("------------------------------")
    List1.AddItem ("Change Dos Prompt")
    List1.AddItem ("------------------------------")
    List1.AddItem ("Show & Edit Date")
    List1.AddItem ("Show & Edit Time")
End Sub

Private Sub List1_Click()
    If List1.Selected(0) = True Then
        Label1.Visible = False
        Text1.Visible = False
    ElseIf List1.Selected(1) = True Then
        Label1.Visible = False
        Text1.Visible = False
    ElseIf List1.Selected(2) = True Then
        Label1.Visible = False
        Text1.Visible = False
    ElseIf List1.Selected(3) = True Then
        Label1.Visible = False
        Text1.Visible = False
    ElseIf List1.Selected(4) = True Then
        Label1.Visible = False
        Text1.Visible = False
    ElseIf List1.Selected(5) = True Then
        Label1.Visible = False
        Text1.Visible = False
    ElseIf List1.Selected(6) = True Then
        Label1.Visible = True
        Text1.Visible = True
        Label1.Caption = "Path and filename:"
    ElseIf List1.Selected(7) = True Then
        Label1.Visible = False
        Text1.Visible = False
    ElseIf List1.Selected(8) = True Then
        Label1.Visible = True
        Text1.Visible = True
        Label1.Caption = "Enter REM Text"
    ElseIf List1.Selected(9) = True Then
        Label1.Visible = True
        Text1.Visible = True
        Label1.Caption = "Enter Text:"
    ElseIf List1.Selected(10) = True Then
        Label1.Visible = False
        Text1.Visible = False
    ElseIf List1.Selected(11) = True Then
        Label1.Visible = False
        Text1.Visible = False
    ElseIf List1.Selected(12) = True Then
        Label1.Visible = False
        Text1.Visible = False
    ElseIf List1.Selected(13) = True Then
        Label1.Visible = False
        Text1.Visible = False
    ElseIf List1.Selected(14) = True Then
        Label1.Visible = True
        Text1.Visible = True
        Label1.Caption = "Enter Pause Text:"
    ElseIf List1.Selected(15) = True Then
        Label1.Visible = True
        Text1.Visible = True
        Label1.Caption = "# of second(s) for delay:"
    ElseIf List1.Selected(17) = True Then
        Label1.Visible = True
        Text1.Visible = True
        Label1.Caption = "Enter file and path:"
    ElseIf List1.Selected(18) = True Then
        Label1.Visible = True
        Text1.Visible = True
        Label1.Caption = "Enter file and path:"
    ElseIf List1.Selected(20) = True Then
        Label1.Visible = True
        Text1.Visible = True
        Label1.Caption = "Enter Directory:"
    ElseIf List1.Selected(21) = True Then
        Label1.Visible = True
        Text1.Visible = True
        Label1.Caption = "Enter Directory:"
    ElseIf List1.Selected(23) = True Then
        Label1.Visible = True
        Text1.Visible = True
        Label1.Caption = "Enter Prompt Text:"
    End If
End Sub

