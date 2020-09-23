VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form4 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Timer & Progress Bar Control"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7755
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
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
      Height          =   975
      Left            =   4200
      Picture         =   "frm_progress.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Pause"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2760
      Picture         =   "frm_progress.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1320
      Picture         =   "frm_progress.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   7080
      Top             =   360
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   1085
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   1320
      Width           =   4455
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Enabled = True

End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
End Sub

Private Sub Command3_Click()
Unload Me
Form2.Show
End Sub

Private Sub Timer1_Timer()
'To add a progress bar to your form got to projects>
'>Components>Microsoft windows Common control 6.0.
'Check this option and ok. You will find the progress bar menu
'added at ur tool box
ProgressBar1.Value = ProgressBar1.Value + 1           'for increament the value of the bar
Label1.Caption = "Percent Done :" & ProgressBar1.Value & " %"    'for display the progress bar value
If ProgressBar1.Value >= 100 Then           'when the bar exceeds the max value
    Timer1.Enabled = False
    ProgressBar1.Value = 0
    Label1.Caption = "Complete"
End If
End Sub
