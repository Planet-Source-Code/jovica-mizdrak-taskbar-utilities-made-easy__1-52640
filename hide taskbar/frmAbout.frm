VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " About"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1110
      Left            =   0
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   1110
      ScaleWidth      =   6225
      TabIndex        =   1
      Top             =   0
      Width           =   6230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   255
      Left            =   5160
      TabIndex        =   0
      Top             =   1440
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "TaskBar Utitlities by Jovica Mizdrak                                            Email: j3d_jovica@hotmail.com"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   4575
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'By Jovica Mizdrak j3d_jovica@hotmail.com
Private Sub Command1_Click()
Me.Hide
frmMain.Show
End Sub
