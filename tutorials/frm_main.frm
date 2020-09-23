VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Common Dialog Control"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdl1 
      Left            =   7680
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
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
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         Picture         =   "frm_main.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2640
         Width           =   4095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Color Dialog"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3120
         TabIndex        =   6
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Help Dialog"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1800
         TabIndex        =   5
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Printer Dialog"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         TabIndex        =   4
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Font Dialog"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3120
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmd_save 
         Caption         =   "Save Dialog"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1800
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmd_open 
         Caption         =   "Open Dialog"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_open_Click()
'Here the name of the common dialog control is cdl1
On Error GoTo nosave            'for tracking the error
cdl1.Filter = "txt files(*.txt)|*.txt|Jpg files(*.jpg)|*.jpg|All files(*.*)|*.*"
cdl1.ShowOpen
'Trapping the error if cancel button is pressed
nosave:
If Err.Number = 32755 Then
    Exit Sub
End If
End Sub

Private Sub Command1_Click()
Unload Me
Form2.Show
End Sub

Private Sub cmd_save_Click()
'Here the name of the common dialog control is cdl1
On Error GoTo nosave            'for tracking the error
cdl1.Filter = "txt files(*.txt)|*.txt|Jpg files(*.jpg)|*.jpg|All files(*.*)|*.*"
cdl1.ShowSave
'Trapping the error if cancel button is pressed
nosave:
If Err.Number = 32755 Then
    Exit Sub
End If
End Sub

Private Sub Command2_Click()
On Error GoTo nofont
'here the constants cdlcfboth means the font dialog wiil show both the
'printer and screen fonts in the dialofg box
cdl1.Flags = cdlCFEffects + cdlCFBoth
cdl1.ShowFont
nofont:
If Err.Number = 32755 Then
    Exit Sub
End If
End Sub

Private Sub Command3_Click()
On Error GoTo noprinter
cdl1.ShowPrinter
noprinter:
If Err.Number = 32755 Then
    Exit Sub
End If
End Sub

Private Sub Command4_Click()
cdl1.ShowHelp
End Sub

Private Sub Command5_Click()
On Error GoTo nocolor
cdl1.Flags = cdlCCRGBInit    'u can also choose the constants cdlccfullopen! Try it
cdl1.ShowColor
nocolor:
If Err.Number = 32755 Then
    Exit Sub
End If

End Sub
