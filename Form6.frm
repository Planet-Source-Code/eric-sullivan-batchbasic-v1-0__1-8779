VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form6"
   ClientHeight    =   3675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6495
   LinkTopic       =   "Form6"
   ScaleHeight     =   3675
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   2160
      Width           =   6495
      Begin VB.Line Line4 
         X1              =   4200
         X2              =   4200
         Y1              =   240
         Y2              =   1200
      End
      Begin VB.Line Line3 
         X1              =   2520
         X2              =   4200
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   2520
         X2              =   2520
         Y1              =   240
         Y2              =   1200
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   2520
         X2              =   4200
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   765
         Left            =   2760
         TabIndex        =   1
         Top             =   240
         Width           =   1185
      End
   End
   Begin VB.Image Image1 
      Height          =   1875
      Left            =   240
      Picture         =   "Form6.frx":0000
      Top             =   120
      Width           =   6000
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   3
      Height          =   3615
      Left            =   0
      Top             =   0
      Width           =   6495
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Call ChangeViews(Line1, Line2, Line3, Line4, False)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.ForeColor = &HFFC0C0
    Call ChangeViews(Line1, Line2, Line3, Line4, False)
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.ForeColor = &HFFC0C0
    Call ChangeViews(Line1, Line2, Line3, Line4, False)
End Sub

Private Sub Label1_Click()
    Form1.Visible = True
    Form6.Visible = False
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.ForeColor = &HFF0000
    Call ChangeViews(Line1, Line2, Line3, Line4, True)
End Sub

Private Sub ChangeViews(L1 As Line, L2 As Line, L3 As Line, L4 As Line, Val As Variant)
    L1.Visible = Val
    L2.Visible = Val
    L3.Visible = Val
    L4.Visible = Val
End Sub
