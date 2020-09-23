VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pop-Up Menu"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7305
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
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
      Left            =   1680
      Picture         =   "frm_pop-ups.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   5535
   End
   Begin VB.Label Label4 
      Caption         =   "Right Click on the Label to see the effect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BAD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "GOOD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HEY, HOW R U???"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
   Begin VB.Menu mnumain 
      Caption         =   "main"
      Visible         =   0   'False
      Begin VB.Menu mnugood 
         Caption         =   "Good"
      End
      Begin VB.Menu mnubad 
         Caption         =   "Bad"
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Form2.Show
End Sub

Private Sub Form_Load()
'for pop-up menu u first make the menu items from the
'menu editor at design time. Keep the main menu visible option=false.
'But u have to make atlesat one submenu visibility property True
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    PopupMenu mnumain
End If
End Sub

Private Sub mnubad_Click()
Label3.Visible = True
Label2.Visible = False
End Sub

Private Sub mnugood_Click()
Label2.Visible = True
Label3.Visible = False
End Sub
