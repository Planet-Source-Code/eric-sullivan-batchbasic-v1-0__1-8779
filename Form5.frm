VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "        Add Commands"
   ClientHeight    =   7290
   ClientLeft      =   13845
   ClientTop       =   3000
   ClientWidth     =   2550
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   2550
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check3 
      Caption         =   "System"
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
      Left            =   240
      TabIndex        =   8
      Top             =   6360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Hidden"
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
      Left            =   240
      TabIndex        =   7
      Top             =   6120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Read Only"
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
      Left            =   240
      TabIndex        =   6
      Top             =   5880
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text2 
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
      Left            =   240
      TabIndex        =   5
      Top             =   6240
      Width           =   2055
   End
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
      Left            =   240
      TabIndex        =   3
      Top             =   5520
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add &Commands"
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
      TabIndex        =   1
      Top             =   6840
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
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
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
      Left            =   240
      TabIndex        =   4
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
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
      Left            =   240
      TabIndex        =   2
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   2400
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   120
      Y1              =   5160
      Y2              =   6720
   End
   Begin VB.Line Line2 
      X1              =   2400
      X2              =   2400
      Y1              =   5160
      Y2              =   6720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   2400
      Y1              =   5160
      Y2              =   5160
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If List1.Selected(0) = True Then
        If Form1.Text1.Text = "" Then
            Form1.Text1.SelText = "copy " & Text1.Text & Text2.Text
        ElseIf Form1.Text1.Text <> "" Then
            Form1.Text1.SelText = vbNewLine & "copy " & Text1.Text & Text2.Text
        End If
    ElseIf List1.Selected(1) = True Then
        If Form1.Text1.Text = "" Then
            Form1.Text1.SelText = "move " & Text1.Text & Text2.Text
        ElseIf Form1.Text1.Text <> "" Then
            Form1.Text1.SelText = vbNewLine & "move " & Text1.Text & Text2.Text
        End If
    ElseIf List1.Selected(2) = True Then
        
        If Check1.Value = Checked Then
            Att1 = "+r"
        ElseIf Check1.Value = Unchecked Then
            Att1 = "-r"
        End If
        
        If Check2.Value = Checked Then
            Att2 = "+h"
        ElseIf Check2.Value = Unchecked Then
            Att2 = "-h"
        End If
        
        If Check3.Value = Checked Then
            Att3 = "+s"
        ElseIf Check3.Value = Unchecked Then
            Att3 = "-s"
        End If
        
        If Form1.Text1.Text = "" Then
            Form1.Text1.SelText = "attrib " & Att1 & Att2 & Att3 & " " & Text1.Text
        ElseIf Form1.Text1.Text <> "" Then
            Form1.Text1.SelText = vbNewLine & "attrib " & Att1 & Att2 & Att3 & " " & Text1.Text
        End If
    ElseIf List1.Selected(3) = True Then
        If Form1.Text1.Text = "" Then
            Form1.Text1.SelText = "ren " & Text1.Text
        ElseIf Form1.Text1.Text <> "" Then
            Form1.Text1.SelText = vbNewLine & "ren " & Text1.Text
        End If
    ElseIf List1.Selected(5) = True Then
        If Form1.Text1.Text = "" Then
            Form1.Text1.SelText = "dir/w"
        ElseIf Form1.Text1.Text <> "" Then
            Form1.Text1.SelText = vbNewLine & "dir/w"
        End If
    ElseIf List1.Selected(6) = True Then
        If Form1.Text1.Text = "" Then
            Form1.Text1.SelText = "dir/p"
        ElseIf Form1.Text1.Text <> "" Then
            Form1.Text1.SelText = vbNewLine & "dir/p"
        End If
    ElseIf List1.Selected(7) = True Then
        If Form1.Text1.Text = "" Then
            Form1.Text1.SelText = "dir"
        ElseIf Form1.Text1.Text <> "" Then
            Form1.Text1.SelText = vbNewLine & "dir"
        End If
    End If
End Sub

Private Sub Form_Load()
    List1.AddItem ("Copy File")
    List1.AddItem ("Move File")
    List1.AddItem ("Customize File Attributes")
    List1.AddItem ("Rename File")
    List1.AddItem ("------------------------------")
    List1.AddItem ("Show Dir List (Width)")
    List1.AddItem ("Show Dir List (Pause)")
    List1.AddItem ("Show Dir List")
End Sub

Private Sub List1_Click()
    If List1.Selected(0) = True Then
        ShowBoxes
        Check1.Visible = False
        Check2.Visible = False
        Check3.Visible = False
        Label1.Caption = "Source File:"
        Label2.Caption = "Destination File:"
    ElseIf List1.Selected(1) = True Then
        ShowBoxes
        Check1.Visible = False
        Check2.Visible = False
        Check3.Visible = False
        Label1.Caption = "Source File:"
        Label2.Caption = "Destination File:"
    ElseIf List1.Selected(2) = True Then
        Label1.Visible = True
        Text1.Visible = True
        Label2.Visible = False
        Text2.Visible = False
        Check1.Visible = True
        Check2.Visible = True
        Check3.Visible = True
        Label1.Caption = "File To Attrib:"
    ElseIf List1.Selected(3) = True Then
        Label1.Visible = True
        Text1.Visible = True
        Check1.Visible = False
        Check2.Visible = False
        Check3.Visible = False
        Label2.Visible = False
        Text2.Visible = False
        Label1.Caption = "Source File To Rename:"
    ElseIf List1.Selected(5) = True Then
        HideAll
    ElseIf List1.Selected(6) = True Then
        HideAll
    ElseIf List1.Selected(7) = True Then
        HideAll
    End If
End Sub

Private Sub HideAll()
    Check1.Visible = False
    Check2.Visible = False
    Check3.Visible = False
    Label2.Visible = False
    Text2.Visible = False
    Label1.Visible = False
    Text1.Visible = False
End Sub

Private Sub ShowBoxes()
    Label1.Visible = True
    Text1.Visible = True
    Label2.Visible = True
    Text2.Visible = True
End Sub

