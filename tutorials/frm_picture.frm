VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Picture Box Control"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7455
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
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
      Left            =   3960
      Picture         =   "frm_picture.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Width           =   3255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save  Picture"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      Top             =   2640
      Width           =   3255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Paste Picture"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   2040
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Copy Picture"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   1440
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Picture"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   840
      Width           =   3255
   End
   Begin VB.PictureBox Picture2 
      Height          =   3375
      Left            =   240
      ScaleHeight     =   3315
      ScaleWidth      =   3075
      TabIndex        =   1
      Top             =   3480
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      Height          =   3255
      Left            =   240
      ScaleHeight     =   3195
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Loads the picture into picture1 control
Picture1.Picture = LoadPicture(App.Path & "\icons\mycat.jpg")
End Sub

Private Sub Command2_Click()
'for copying any item to clipboard
'first clear the clipboard
Clipboard.Clear
'Now coy the picture
Clipboard.SetData Picture1.Picture
End Sub

Private Sub Command3_Click()
'paste the picture.Here I am not using the Loadpicture method
'because the picture already exists at the clipboard
Picture2.Picture = Clipboard.GetData
End Sub

Private Sub Command4_Click()
'Saves the picture
SavePicture Picture2.Picture, (App.Path & "\save\mycat.bmp")
End Sub

Private Sub Command5_Click()
Unload Me
Form2.Show
End Sub
