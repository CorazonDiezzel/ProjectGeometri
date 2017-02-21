VERSION 5.00
Begin VB.Form Form_Menu 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10485
   DrawMode        =   16  'Merge Pen
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Form_Menu.frx":0000
   ScaleHeight     =   6990
   ScaleWidth      =   10485
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      DisabledPicture =   "Form_Menu.frx":1BA07
      DownPicture     =   "Form_Menu.frx":2127B
      Height          =   2535
      Index           =   0
      Left            =   600
      MaskColor       =   &H8000000A&
      Picture         =   "Form_Menu.frx":26AEF
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      Width           =   4095
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400000&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   10515
      TabIndex        =   0
      Top             =   0
      Width           =   10575
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "MENU"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   855
         Index           =   0
         Left            =   3840
         TabIndex        =   1
         Top             =   120
         Width           =   2895
      End
      Begin VB.Image Image1 
         Height          =   975
         Left            =   9360
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Image cubic_grey 
      Appearance      =   0  'Flat
      Height          =   840
      Left            =   8640
      Picture         =   "Form_Menu.frx":2C363
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   840
   End
   Begin VB.Image cubic_blue 
      Height          =   960
      Left            =   5520
      Picture         =   "Form_Menu.frx":2F446
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   960
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   3960
      Left            =   5520
      Picture         =   "Form_Menu.frx":30954
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   3960
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "BANGUN RUANG"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Index           =   1
      Left            =   5760
      TabIndex        =   4
      Top             =   2040
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "BANGUN DATAR"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Index           =   0
      Left            =   840
      TabIndex        =   3
      Top             =   2040
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MENU"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   855
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "Form_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = cubic_grey.Picture
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = cubic_blue.Picture
End Sub
