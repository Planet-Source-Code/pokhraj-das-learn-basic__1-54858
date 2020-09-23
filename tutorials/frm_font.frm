VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Font Example"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5700
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   975
      Left            =   240
      Picture         =   "frm_font.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5280
      Width           =   5055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "BLUE"
      Height          =   255
      Left            =   3960
      TabIndex        =   12
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GREEN"
      Height          =   255
      Left            =   3960
      TabIndex        =   11
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "RED"
      Height          =   255
      Left            =   3960
      TabIndex        =   10
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   240
      TabIndex        =   9
      Top             =   3240
      Width           =   5295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose Options"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1215
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   3135
      Begin VB.CheckBox Check4 
         Caption         =   "Italics"
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
         Left            =   2160
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Bold"
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
         Left            =   2160
         TabIndex        =   7
         Top             =   840
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         Caption         =   "UnderLine"
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
         TabIndex        =   6
         Top             =   840
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "StrikeThrough"
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
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   720
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Sample"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   13
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Fonts Size"
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
      Index           =   1
      Left            =   3240
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Fonts Name"
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
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = vbChecked Then
    Text1.Font.Strikethrough = True
ElseIf Check1.Value = False Then
    Text1.Font.Strikethrough = False
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = vbChecked Then
    Text1.Font.Underline = True
ElseIf Check2.Value = False Then
    Text1.Font.Underline = False
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = vbChecked Then
    Text1.Font.Bold = True
ElseIf Check3.Value = False Then
    Text1.Font.Bold = False
End If
End Sub

Private Sub Check4_Click()
If Check4.Value = vbChecked Then
    Text1.Font.Italic = True
ElseIf Check1.Value = False Then
    Text1.Font.Italic = False
End If
End Sub

Private Sub Combo1_Click()
Text1.Font.Name = Combo1.Text
End Sub

Private Sub Combo2_Click()
Text1.Font.Size = Combo2.Text
End Sub

Private Sub Command1_Click()
Text1.ForeColor = vbRed
End Sub

Private Sub Command2_Click()
Text1.ForeColor = vbGreen
End Sub

Private Sub Command3_Click()
Text1.ForeColor = vbBlue
End Sub

Private Sub Command4_Click()
Unload Me
Form2.Show
End Sub

Private Sub Form_Load()
'for adding the fonts name
Dim i As Integer, j As Integer
For i = 1 To Screen.FontCount - 1
    Combo1.AddItem Screen.Fonts(i)
Next i
'for adding the size
For j = 8 To 40 Step 2
    Combo2.AddItem j
Next j
End Sub
