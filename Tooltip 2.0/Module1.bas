Attribute VB_Name = "Module1"
Global Top1 As Integer
Global Left1 As Integer
Global X1 As Integer
Global Y1 As Integer
Global Start As Boolean
Global P1 As Integer
Global S1 As Integer
Global Word As String
Global Word1 As String
Global Word2 As String
Global Word3 As String
Public Sub RealLabel()
     With Form2
     .Label2.AutoSize = True
     .Picture1.AutoSize = True
     .Width = .Label2.Width + .Picture1.Width
     .Height = .Picture1.Height + 23
     .Shape1.Height = .Height
     .Shape1.Width = .Width
     .Show
     End With
End Sub
Public Sub Tooltip()
      Form2.Top = Y1 + Top1 + Form1.Top + 300
      Form2.Left = X1 + Left1 + Form1.Left - 700
End Sub
Public Sub ScrollText()
       Top1 = Form1.Text1.Top: Left1 = Form1.Text1.Left
       Call Module1.Tooltip
       With Form2
       .Timer1.Enabled = True
       ' If too slow change interval...
       .Timer1.Interval = 80
       .Width = .Text1.Width + 20
       .Height = .Text1.Height + 20
       .Picture1.Visible = False
       .Picture2.Visible = False
       .Label2.Visible = False
       .Show
       End With
       Word1 = String(20, " ") + Word1
       S1 = 2
       Start = False
End Sub
Public Sub PictureTooltip()
       Top1 = Form1.Text2.Top
       Left1 = Form1.Text2.Left
       With Form2
       .Timer1.Enabled = False
       .Text1.Visible = False
       .Picture1.Width = 1890
       End With
       Call Module1.Tooltip
       Call Module1.RealLabel
       Form2.Show
       Start = False
End Sub
Public Sub Animation()
     Top1 = Form1.Text3.Top
     Left1 = Form1.Text3.Left
     With Form2
     .Timer1.Enabled = True
     ' if too fast change interval...
     .Timer1.Interval = 150
     .Label2.FontSize = 10
     .Label2.FontName = "Bookman Old Style"
     .Picture2.Visible = False
     .Text1.Visible = False
     .Label2.FontItalic = True
     .Label2.ForeColor = vbBlue
     .Picture1.Width = 500
     End With
     Call Module1.Tooltip
     S1 = 1
     Start = False
End Sub
