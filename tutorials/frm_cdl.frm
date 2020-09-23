VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Main form "
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5760
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command10 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      Picture         =   "frm_cdl.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3240
      Width           =   5535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Common Dialog Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose Your Option"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.CommandButton Command9 
         Caption         =   "Slider Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Message Box"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Picture Box Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Date/Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Pop-Up Menu"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Fonts"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Timer Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "List Box Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'unloading main form
Unload Me
'opening the new form
Form1.Show

End Sub

Private Sub Command10_Click()
End
End Sub

Private Sub Command2_Click()
Unload Me
Form3.Show
End Sub

Private Sub Command3_Click()
Unload Me
Form4.Show
End Sub

Private Sub Command4_Click()
Unload Me
Form5.Show
End Sub

Private Sub Command5_Click()
Unload Me
Form6.Show
End Sub

Private Sub Command6_Click()
Form7.Show
Unload Me
End Sub

Private Sub Command7_Click()
Unload Me
Form8.Show
End Sub

Private Sub Command8_Click()
Unload Me
Form10.Show
End Sub

Private Sub Command9_Click()
Unload Me
Form9.Show
End Sub
