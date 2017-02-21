VERSION 5.00
Begin VB.Form Form_Login 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10035
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   10035
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox forminput_password 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1680
      TabIndex        =   2
      Text            =   "Password"
      Top             =   3000
      Width           =   3975
   End
   Begin VB.TextBox forminput_username 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1680
      TabIndex        =   1
      Text            =   "Username"
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Image login_default 
      Height          =   540
      Left            =   8400
      Picture         =   "Login.frx":0000
      Stretch         =   -1  'True
      Top             =   2040
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Image login_hover 
      Height          =   540
      Left            =   8400
      Picture         =   "Login.frx":2015
      Stretch         =   -1  'True
      Top             =   2640
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Image login_button 
      Height          =   2460
      Left            =   6000
      Picture         =   "Login.frx":6511
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   2685
   End
   Begin VB.Image exit_deactive 
      Height          =   1320
      Left            =   7440
      Picture         =   "Login.frx":8526
      Stretch         =   -1  'True
      Top             =   3720
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image exit_hover 
      Height          =   1320
      Left            =   8040
      Picture         =   "Login.frx":9284
      Stretch         =   -1  'True
      Top             =   3720
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   4
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   2880
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   4
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Image icon_password 
      Appearance      =   0  'Flat
      Height          =   885
      Left            =   480
      Picture         =   "Login.frx":9FE2
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   885
   End
   Begin VB.Image icon_user 
      Appearance      =   0  'Flat
      Height          =   1245
      Left            =   240
      Picture         =   "Login.frx":A773
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1365
   End
   Begin VB.Image exit_button 
      Height          =   960
      Left            =   8880
      Picture         =   "Login.frx":AC74
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sign In"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9975
   End
   Begin VB.Image Image1 
      Height          =   5415
      Left            =   0
      Picture         =   "Login.frx":B9D2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10005
   End
End
Attribute VB_Name = "Form_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim user_name, user_pass As String
Private Sub exit_button_Click()
Unload Me
End Sub

Private Sub exit_button_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
exit_button.Picture = exit_hover.Picture
End Sub

Private Sub forminput_password_Click()
forminput_password.Text = ""
End Sub

Private Sub forminput_username_Click()
forminput_username.Text = ""
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If exit_button.Picture = exit_hover.Picture Then
    exit_button.Picture = exit_deactive.Picture
Else
End If
If login_button.Picture = login_hover.Picture Then
    login_button.Picture = login_default.Picture
Else
End If

End Sub


Private Sub login_button_Click()
If user_name = "" And user_pass = "" Then
user_name = "admin"
user_pass = "admin"
Else
End If
If (forminput_username.Text) = user_name And (forminput_password.Text) = user_pass Then
MsgBox ("Login Berhasil")
Me.Hide
Form_Menu.Show
Else
MsgBox ("Login Gagal")
End If
End Sub

Private Sub login_button_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
login_button.Picture = login_hover.Picture
End Sub


