VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "List Box control"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7845
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   7845
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
      Left            =   1080
      Picture         =   "frm_list.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6240
      Width           =   5415
   End
   Begin VB.Frame Frame3 
      Caption         =   "Choose Your Option"
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
      Height          =   1455
      Left            =   4560
      TabIndex        =   9
      Top             =   4680
      Width           =   3135
      Begin VB.CommandButton cmd_clr 
         Caption         =   "Clear List"
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
         Left            =   360
         TabIndex        =   12
         Top             =   1080
         Width           =   2415
      End
      Begin VB.CommandButton cmd_rem 
         Caption         =   "Remove Selected Item"
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
         Left            =   360
         TabIndex        =   11
         Top             =   720
         Width           =   2415
      End
      Begin VB.CommandButton cmd_addnew 
         Caption         =   "Add New Element"
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
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Choose Your Option"
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
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   4680
      Width           =   3015
      Begin VB.CommandButton cmd_clear 
         Caption         =   "Clear List"
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
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   2535
      End
      Begin VB.CommandButton cmd_add 
         Caption         =   "Add New Element"
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
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2535
      End
      Begin VB.CommandButton cmd_Remove 
         Caption         =   "Remove selected Item"
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
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "List Demo"
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.CommandButton cmd_trans 
         Caption         =   "<"
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
         Left            =   3360
         TabIndex        =   4
         Top             =   2280
         Width           =   855
      End
      Begin VB.CommandButton cmd_transfer 
         Caption         =   ">"
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
         Left            =   3360
         TabIndex        =   3
         Top             =   1680
         Width           =   855
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3420
         Left            =   4440
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   2895
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3420
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   2895
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_add_Click()
Dim strinput As String
strinput = InputBox("Enter Item To Add to List :", "Add")
List1.AddItem strinput
End Sub

Private Sub cmd_addnew_Click()
'For adding the new item in listbox2
Dim strinput As String
strinput = InputBox("Enter Item To Add to List :", "Add")
List2.AddItem strinput
End Sub

Private Sub cmd_clear_Click()
'The clear method clears the entire list box.
List1.Clear
End Sub

Private Sub cmd_clr_Click()
'clear the items of list2
List2.Clear
End Sub

Private Sub cmd_rem_Click()
If List2.ListIndex >= 0 Then
List2.RemoveItem List2.ListIndex
Else
MsgBox "Select Item for remove", vbInformation, "Error"
End If
End Sub

Private Sub cmd_Remove_Click()
'The removeitem property removes the item
'from the listbox which is particularly selected.If no item
' is being selected an error occured
If List1.ListIndex >= 0 Then
List1.RemoveItem List1.ListIndex
Else
    MsgBox "Select Item for remove", vbInformation, "Error"
End If
End Sub


Private Sub cmd_trans_Click()
If List2.ListIndex >= 0 Then
List1.AddItem List2.Text
List2.RemoveItem List2.ListIndex
End If
End Sub

Private Sub cmd_transfer_Click()
If List1.ListIndex >= 0 Then
List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
End If
End Sub

Private Sub Command1_Click()
Unload Me
Form2.Show
End Sub
