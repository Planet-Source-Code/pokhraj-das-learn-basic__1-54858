VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form9 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Slider Control"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6825
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   360
      Picture         =   "frm_slider.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   6375
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   1296
      _Version        =   393216
      Appearance      =   1
      Max             =   25
      Scrolling       =   1
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   675
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   1191
      _Version        =   393216
      LargeChange     =   1
      Max             =   25
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
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
      Left            =   1320
      TabIndex        =   2
      Top             =   960
      Width           =   4335
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Form2.Show
End Sub

Private Sub Slider1_Change()
ProgressBar1.Value = Slider1.Value
Label1.Caption = "The value of slider is now: " & Slider1.Value
If ProgressBar1.Value >= 25 Then
    Exit Sub
End If
End Sub

Private Sub Slider1_Click()
ProgressBar1.Value = Slider1.Value
Label1.Caption = "The value of slider is now: " & Slider1.Value
If ProgressBar1.Value >= 25 Then
    Exit Sub
End If
End Sub



Private Sub Slider1_Scroll()
ProgressBar1.Value = Slider1.Value
Label1.Caption = "The value of slider is now: " & Slider1.Value
If ProgressBar1.Value >= 25 Then
    Exit Sub
End If
End Sub
