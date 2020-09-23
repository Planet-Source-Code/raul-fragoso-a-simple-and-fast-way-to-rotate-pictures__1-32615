VERSION 5.00
Begin VB.Form frmRotate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rotate"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDest 
      AutoRedraw      =   -1  'True
      Height          =   6705
      Left            =   4560
      ScaleHeight     =   443
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   286
      TabIndex        =   4
      Top             =   120
      Width           =   4350
      Begin VB.Shape shSel 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   2145
         Shape           =   3  'Circle
         Top             =   3330
         Width           =   135
      End
   End
   Begin VB.Timer tmrAuto 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4560
      Top             =   6960
   End
   Begin VB.CommandButton cmdAuto 
      Caption         =   "Automatic"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton cmdRotRight 
      Caption         =   "Rotate Right"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton cmdRotLeft 
      Caption         =   "Rotate Left"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   6960
      Width           =   1335
   End
   Begin VB.PictureBox picSource 
      Height          =   6705
      Left            =   120
      Picture         =   "frmRotate.frx":0000
      ScaleHeight     =   443
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   286
      TabIndex        =   0
      Top             =   120
      Width           =   4350
   End
End
Attribute VB_Name = "frmRotate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_Angle As Long

Private Sub cmdAuto_Click()

    If cmdAuto.Caption = "Automatic" Then
        cmdAuto.Caption = "Stop"
        tmrAuto.Enabled = True
    Else
        cmdAuto.Caption = "Automatic"
        tmrAuto.Enabled = False
    End If

End Sub

Private Sub cmdRotLeft_Click()

    tmrAuto.Enabled = False
    m_Angle = m_Angle - 5
    If m_Angle < -360 Then m_Angle = 0
    
    picDest.Cls
    RotateDC picDest.hdc, shSel.Left + (shSel.Width / 2), shSel.Top + (shSel.Height / 2), picSource.hdc, picSource.Picture.Handle, m_Angle

End Sub

Private Sub cmdRotRight_Click()

    tmrAuto.Enabled = False
    m_Angle = m_Angle + 5
    If m_Angle > 360 Then m_Angle = 0
    
    picDest.Cls
    RotateDC picDest.hdc, shSel.Left + (shSel.Width / 2), shSel.Top + (shSel.Height / 2), picSource.hdc, picSource.Picture.Handle, m_Angle

End Sub

Private Sub Form_Load()

    shSel.Left = (picDest.ScaleWidth / 2) - (shSel.Width / 2)
    shSel.Top = (picDest.ScaleHeight / 2) - (shSel.Height / 2)

End Sub

Private Sub picDest_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    shSel.Move x - (shSel.Width / 2), y - (shSel.Height / 2)

End Sub

Private Sub picDest_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then
        shSel.Move x - (shSel.Width / 2), y - (shSel.Height / 2)
    End If

End Sub

Private Sub tmrAuto_Timer()

    m_Angle = m_Angle + 5
    If m_Angle > 360 Then m_Angle = 0
    picDest.Cls
    RotateDC picDest.hdc, shSel.Left + (shSel.Width / 2), shSel.Top + (shSel.Height / 2), picSource.hdc, picSource.Picture.Handle, m_Angle

End Sub
