VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BatchBasic v1.0 [HELP]"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8145
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label9 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2520
      TabIndex        =   8
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "dir/switch filter( *.exe)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   4215
   End
   Begin VB.Label Label7 
      Caption         =   "Switches that can be used to determine different ways of listing the directories:               /b /c /ch /l /s /p /w"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   6
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Line Line4 
      X1              =   4440
      X2              =   120
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   120
      Y1              =   1200
      Y2              =   2640
   End
   Begin VB.Line Line2 
      X1              =   4440
      X2              =   4440
      Y1              =   1200
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   4440
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "BatchBasic v1.0 By: Eric Sullivan"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   4680
      TabIndex        =   4
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   $"Form7.frx":0000
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   3
      Top             =   360
      Width           =   5175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "PROMPT "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "will enable or disable the ability to make dos show the commands"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "@echo on/off: "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Label5.Caption = "$Q = (equal sign)" & _
    vbNewLine & "$$ $ (dollar sign)" & _
    vbNewLine & "$T Current time" & _
    vbNewLine & "$D Current date" & _
    vbNewLine & "$P Current drive and path" & _
    vbNewLine & "$V Windows version number" & _
    vbNewLine & "$N Current drive" & _
    vbNewLine & "$G > (greater-than sign)" & _
    vbNewLine & "$L & (less-than sign)" & _
    vbNewLine & "$B | (pipe)" & _
    vbNewLine & "$H Backspace (erases previous character)" & _
    vbNewLine & "$E  Escape code (ASCII code 27)" & _
    vbNewLine & "$_ Carriage return and linefeed"
    Label9.Caption = "You can also add filter's (*.*, *.exe) to only view cirtain files"
End Sub

