VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "Form1.frx":0000
      Top             =   3000
      Width           =   3735
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   3960
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form1.frx":005E
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   1080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "Form1.frx":00A6
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   0
      X2              =   6600
      Y1              =   320
      Y2              =   320
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   5295
      Left            =   0
      Top             =   0
      Width           =   6600
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   280
      Left            =   5760
      MouseIcon       =   "Form1.frx":00E6
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   30
      Width           =   675
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tooltip Picture 2.0 By Nobody  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   320
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub Command1_Click()
       MsgBox "Don't forget for vote!", , "Tooltip By Nobody"
       End
End Sub

Private Sub Form_Load()
        S1 = 0
        P1 = 0
        Start = True
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Unload Form2
        Start = True
        A = 1
End Sub
Private Sub Label2_Click()
        MsgBox "Don't forget for vote!", , "Tooltip By Nobody"
        End
End Sub
Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       If Start = True Then
       X1 = X: Y1 = Y
       Word1 = "You never see this before." & _
       " Tooltip with scrooltext. Don't forget for vote." & _
       " If you have some suggestion on this programm, send me" & _
       " e-mail or leave on Planet... Today is: " & Date & _
       ". Time: " & Time
       Call Module1.ScrollText
       End If
End Sub

Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       If Start = True Then
       X1 = X: Y1 = Y
       Word = "Any comments please" _
       & vbCrLf & "send.VOTE for my work!"
       Call Module1.PictureTooltip
       End If
End Sub

Private Sub Text3_Click()
        S1 = 0
        Unload Form2
End Sub

Private Sub Text3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       If Start = True Then
       X1 = X: Y1 = Y
       Word2 = " E - m a i l  :                  " _
       & vbCrLf & " ssurlina@inet.hr !"
       Word3 = " Please vote and leave   " _
       & vbCrLf & "comments on this code !"
       Call Module1.Animation
       End If
End Sub
