VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2985
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5520
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2640
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   4800
      Top             =   360
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Height          =   2955
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   5505
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "If my Program helps any one please Comment. "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2040
         TabIndex        =   3
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Developed By Pokhraj Das."
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
         Left            =   1680
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
      Begin VB.Image imgLogo 
         Height          =   1785
         Left            =   240
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 1
If ProgressBar1.Value >= 100 Then
    Unload Me
    Timer1.Enabled = False
    Form2.Show
End If
End Sub
