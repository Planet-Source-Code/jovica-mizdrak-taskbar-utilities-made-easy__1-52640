VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " TaskBar Utilities"
   ClientHeight    =   2520
   ClientLeft      =   30
   ClientTop       =   195
   ClientWidth     =   3360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   3360
   StartUpPosition =   3  'Windows Default
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command1 
      Caption         =   "Apply"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   972
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   2160
      Width           =   972
   End
   Begin VB.CommandButton Command2 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   2160
      Width           =   972
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Tray Icons"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   1800
      TabIndex        =   9
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Tray Button"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   8
      Top             =   840
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Tray Clock"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   1800
      TabIndex        =   7
      Top             =   480
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Task Programs"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "System Tray"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Tool Bars"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Start Button"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Whole TaskBar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Hide Selected Items:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   120
      X2              =   3240
      Y1              =   1560
      Y2              =   1560
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'By Jovica Mizdrak j3d_jovica@hotmail.com

Private Sub Check1_Click(Index As Integer)
Dim tray As Integer

Select Case Index
    Case 0
    
    If Check1(0).Value = 1 Then
    For tray = 5 To 7
    Check1(tray).Value = 1
    Check1(tray).Enabled = False
    Next tray
    ElseIf Check1(0).Value = 0 Then
    For tray = 5 To 7 Step 1
    Check1(tray).Value = 0
    Check1(tray).Enabled = True
    Next tray
    End If
    Case 4
    If Check1(4).Value = 1 Then
    Check1(0).Value = 0
    Check1(0).Enabled = False
    Check1(1).Value = 0
    Check1(1).Enabled = False
    Check1(2).Value = 0
    Check1(2).Enabled = False
    Check1(3).Value = 0
    Check1(3).Enabled = False
    Check1(5).Value = 0
    Check1(5).Enabled = False
    Check1(6).Value = 0
    Check1(6).Enabled = False
    Check1(7).Value = 0
    Check1(7).Enabled = False
    ElseIf Check1(4).Value = 0 Then
    Check1(0).Value = 0
    Check1(0).Enabled = True
    Check1(1).Value = 0
    Check1(1).Enabled = True
    Check1(2).Value = 0
    Check1(2).Enabled = True
    Check1(3).Value = 0
    Check1(3).Enabled = True
    Check1(5).Value = 0
    Check1(5).Enabled = True
    Check1(6).Value = 0
    Check1(6).Enabled = True
    Check1(7).Value = 0
    Check1(7).Enabled = True
    End If
    
End Select
End Sub

Private Sub Command1_Click()
'By Jovica Mizdrak j3d_jovica@hotmail.com
Dim hnd As Long
Dim toolbar, tray, prog, clock, traybutton, trayicon, OurHandle As Long

'hide the taskbar
hnd = FindWindow("Shell_traywnd", "") 'get the Window

toolbar = FindWindowEx(hnd, 0, "ReBarWindow32", vbNullString)
prog = FindWindowEx(toolbar, 0, "MSTaskSwWClass", vbNullString) 'Programs
toolbar = FindWindowEx(toolbar, 0, "ToolBarWindow32", vbNullString) 'Toolbar

tray = FindWindowEx(hnd, 0, "TrayNotifyWnd", vbNullString) 'Tray

OurHandle = FindWindowEx(hnd, 0, "Button", vbNullString) ' Start Button

clock = FindWindowEx(hnd, 0, "TrayNotifyWnd", vbNullString)
clock = FindWindowEx(clock, 0, "TrayClockWClass", vbNullString) 'Toolbar like: Quick lunch

traybutton = FindWindowEx(hnd, 0, "TrayNotifyWnd", vbNullString)
traybutton = FindWindowEx(traybutton, 0, "Button", vbNullString) ' Tray Button

trayicon = FindWindowEx(hnd, 0, "TrayNotifyWnd", vbNullString)
trayicon = FindWindowEx(trayicon, 0, "SysPager", vbNullString)
trayicon = FindWindowEx(trayicon, 0, "ToolBarWindow32", vbNullString) ' Tray Icons


If Check1(0).Value = 1 Then
ShowWindow tray, 0
Else
ShowWindow tray, 5
End If

If Check1(1).Value = 1 Then
ShowWindow prog, 0
Else
ShowWindow prog, 5
End If

If Check1(2).Value = 1 Then
ShowWindow toolbar, 0
Else
ShowWindow toolbar, 5
End If

If Check1(3).Value = 1 Then
ShowWindow OurHandle, 0
Else
ShowWindow OurHandle, 5
End If

If Check1(4).Value = 1 Then
ShowWindow hnd, 0
Else
ShowWindow hnd, 5
End If

If Check1(5).Value = 1 Then
ShowWindow clock, 0
Else
ShowWindow clock, 5
End If

If Check1(6).Value = 1 Then
ShowWindow traybutton, 0
Else
ShowWindow traybutton, 5
End If

If Check1(7).Value = 1 Then
ShowWindow trayicon, 0
Else
ShowWindow trayicon, 5
End If

saveSettings

End Sub

Private Sub Command2_Click()
frmMain.Hide
frmAbout.Show
End Sub

Private Sub Command3_Click()
Unload Me
End
End Sub

Private Sub Form_Initialize()
    If App.PrevInstance = True Then
        End
    End If
End Sub

Private Sub Form_Load()
Check1(0).Value = GetSetting(App.EXEName, "Settings", "Check1(0)", 0)
Check1(1).Value = GetSetting(App.EXEName, "Settings", "Check1(1)", 0)
Check1(2).Value = GetSetting(App.EXEName, "Settings", "Check1(2)", 0)
Check1(3).Value = GetSetting(App.EXEName, "Settings", "Check1(3)", 0)
Check1(4).Value = GetSetting(App.EXEName, "Settings", "Check1(4)", 0)
Check1(5).Value = GetSetting(App.EXEName, "Settings", "Check1(5)", 0)
Check1(6).Value = GetSetting(App.EXEName, "Settings", "Check1(6)", 0)
Check1(7).Value = GetSetting(App.EXEName, "Settings", "Check1(7)", 0)
End Sub

Private Sub saveSettings()
SaveSetting App.EXEName, "Settings", "Check1(0)", Check1(0).Value
SaveSetting App.EXEName, "Settings", "Check1(1)", Check1(1).Value
SaveSetting App.EXEName, "Settings", "Check1(2)", Check1(2).Value
SaveSetting App.EXEName, "Settings", "Check1(3)", Check1(3).Value
SaveSetting App.EXEName, "Settings", "Check1(4)", Check1(4).Value
SaveSetting App.EXEName, "Settings", "Check1(5)", Check1(5).Value
SaveSetting App.EXEName, "Settings", "Check1(6)", Check1(6).Value
SaveSetting App.EXEName, "Settings", "Check1(7)", Check1(7).Value
End Sub

'By Jovica Mizdrak j3d_jovica@hotmail.com
