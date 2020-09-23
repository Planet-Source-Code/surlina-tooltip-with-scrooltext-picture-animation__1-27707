VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3345
   LinkTopic       =   "Form2"
   ScaleHeight     =   4995
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   90
      Left            =   2760
      Top             =   2400
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Top             =   1680
      Width           =   135
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   5
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   465
      ScaleWidth      =   1890
      TabIndex        =   4
      Top             =   10
      Width           =   1920
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9
         Charset         =   238
         Weight          =   300
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Left            =   10
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   " "
      Top             =   10
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      DrawWidth       =   3
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   10
      ScaleHeight     =   495
      ScaleWidth      =   3240
      TabIndex        =   2
      Top             =   10
      Width           =   3240
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   495
      TabIndex        =   3
      Top             =   20
      Width           =   45
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   11
      Left            =   1560
      Picture         =   "Form2.frx":1415
      Top             =   4200
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   10
      Left            =   1560
      Picture         =   "Form2.frx":1885
      Top             =   3600
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   9
      Left            =   1680
      Picture         =   "Form2.frx":1CFF
      Top             =   3000
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   8
      Left            =   1560
      Picture         =   "Form2.frx":2169
      Top             =   2400
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   7
      Left            =   1560
      Picture         =   "Form2.frx":25DD
      Top             =   1800
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   6
      Left            =   1560
      Picture         =   "Form2.frx":2A56
      Top             =   1080
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   5
      Left            =   240
      Picture         =   "Form2.frx":2EEF
      Top             =   4200
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   4
      Left            =   240
      Picture         =   "Form2.frx":337B
      Top             =   3600
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   3
      Left            =   240
      Picture         =   "Form2.frx":37D1
      Top             =   2880
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   240
      Picture         =   "Form2.frx":3C28
      Top             =   2400
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   240
      Picture         =   "Form2.frx":40A2
      Top             =   1680
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   360
      Picture         =   "Form2.frx":4512
      Top             =   1080
      Width           =   480
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Picture2.Print Word
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Unload Me
    Start = True
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Unload Me
    Start = True
End Sub
Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Unload Me
    Start = True
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Unload Me
    Start = True
End Sub

Private Sub Timer1_Timer()
     If Start = False Then
     Select Case S1
        Case 1
           If P1 = 5 Then
           Label2.Caption = ""
           Label2.Caption = Label2.Caption & Word2
           Call Module1.RealLabel
           End If
           If P1 = 0 Then
           Label2.Caption = ""
           Label2.Caption = Label2.Caption & Word3
           Call Module1.RealLabel
           End If
           Form2.Picture1.Picture = Form2.Image1(P1).Picture
           P1 = P1 + 1
           If P1 > 11 Then P1 = 0
        Case 2
           Word1 = Mid(Word1, 2) & Left(Word1, 1)
           Text1.Text = Word1
           End Select
           End If
End Sub
