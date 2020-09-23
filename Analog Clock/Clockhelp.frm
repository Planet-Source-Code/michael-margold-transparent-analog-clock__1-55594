VERSION 5.00
Begin VB.Form HelpFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   Icon            =   "Clockhelp.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Clockhelp.frx":0742
   ScaleHeight     =   5790
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   180
      Left            =   1995
      TabIndex        =   2
      Top             =   4935
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   330
      Left            =   2355
      Picture         =   "Clockhelp.frx":1461
      TabIndex        =   0
      Top             =   5355
      Width           =   1155
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Created by Michael Margold"
      Height          =   240
      Left            =   60
      TabIndex        =   5
      Top             =   4230
      Width           =   5760
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Transparent Analog Clock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   60
      TabIndex        =   4
      Top             =   3735
      Width           =   5760
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "     Show Help On Start Up"
      Height          =   240
      Left            =   60
      TabIndex        =   3
      Top             =   4935
      Width           =   5760
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "http://www.soft-collection.com"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   60
      MouseIcon       =   "Clockhelp.frx":2263
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   4575
      Width           =   5760
   End
End
Attribute VB_Name = "HelpFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
  MainFrm.ShowHelp = Check1.Value
  MainFrm.SaveFile
End Sub
Private Sub Command1_Click()
  Me.Hide
End Sub
Private Sub Label1_Click()
  Call Link(Label1.Caption)
End Sub
