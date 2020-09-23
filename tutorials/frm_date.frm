VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Date And Time"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5160
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5160
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
      Height          =   915
      Left            =   120
      Picture         =   "frm_date.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   4935
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1080
      Top             =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4935
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Form2.Show
End Sub

Private Sub Form_Load()
Label1.Caption = "Today is : " & Format(Now(), "Long Date")
End Sub

Private Sub Timer1_Timer()
Label2.Caption = "Now it is :-" & Format(Now(), "Long time")
End Sub
