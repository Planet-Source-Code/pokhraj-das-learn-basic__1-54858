VERSION 5.00
Begin VB.Form Form10 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Message Box"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8760
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Picture         =   "frm_message.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3240
      Width           =   5415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click For The Button Effect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   16
      Top             =   2520
      Width           =   5415
   End
   Begin VB.Frame Frame2 
      Caption         =   "Choose For Buttons"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2295
      Left            =   6120
      TabIndex        =   10
      Top             =   2400
      Width           =   2535
      Begin VB.OptionButton Option9 
         Caption         =   "Abort Rety Ignore"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   2055
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Yes No Cancel "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   1695
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Yes No "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton Option6 
         Caption         =   "OK Cancel  "
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton Option5 
         Caption         =   "OK "
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Click For The Style Effect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   5415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose Style"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2055
      Left            =   6120
      TabIndex        =   2
      Top             =   240
      Width           =   2535
      Begin VB.OptionButton Option4 
         Caption         =   "Critical"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Question"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Exclamation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      TabIndex        =   1
      Text            =   " "
      Top             =   240
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Text            =   " "
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Enter Text for Prompt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Enter For Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   1200
      Width           =   2415
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strmsg As String
Private Sub Command1_Click()
If Option5.Value = True Then
    strmsg = MsgBox(Text2.Text, vbOKOnly, Text1.Text)
ElseIf Option6.Value = True Then
    strmsg = MsgBox(Text2.Text, vbOKCancel, Text1.Text)
ElseIf Option7.Value = True Then
    strmsg = MsgBox(Text2.Text, vbYesNo, Text1.Text)
ElseIf Option8.Value = True Then
    strmsg = MsgBox(Text2.Text, vbYesNoCancel, Text1.Text)
ElseIf Option9.Value = True Then
    strmsg = MsgBox(Text2.Text, vbAbortRetryIgnore, Text1.Text)
End If
End Sub

Private Sub Command2_Click()
Unload Me
Form2.Show
End Sub

Private Sub Command3_Click()

If Option1.Value = True Then
strmsg = MsgBox(Text2.Text, vbExclamation, Text1.Text)
ElseIf Option2.Value = True Then
    strmsg = MsgBox(Text2.Text, vbQuestion, Text1.Text)
ElseIf Option3.Value = True Then
    strmsg = MsgBox(Text2.Text, vbInformation, Text1.Text)
ElseIf Option4.Value = True Then
    strmsg = MsgBox(Text2.Text, vbCritical, Text1.Text)
End If

End Sub

