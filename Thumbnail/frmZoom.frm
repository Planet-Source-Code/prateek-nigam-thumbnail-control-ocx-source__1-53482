VERSION 5.00
Begin VB.Form frmZoom 
   Caption         =   "Form1"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   9255
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1470
      Left            =   3660
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1950
   End
End
Attribute VB_Name = "frmZoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim incre As Double
    incre = 0.01
    If KeyAscii = 45 Then
        Image1.Visible = False
        Image1.Height = Image1.Height - (Image1.Height * incre)
        Image1.Width = Image1.Width - (Image1.Width * incre)
        Image1.Visible = True
    End If
    If KeyAscii = 43 Then
        Image1.Visible = False
        Image1.Height = Image1.Height + (Image1.Height * incre)
        Image1.Width = Image1.Width + (Image1.Width * incre)
        Image1.Visible = True
    End If
    SetImageInCenter
End Sub

Private Sub Form_Resize()
    On Error GoTo errh
    If Me.WindowState = vbMinimized Then Exit Sub
    SetImageInCenter
errh:
End Sub
Private Sub SetImageInCenter()
    Image1.Move (ScaleWidth - Image1.Width) / 2, (ScaleHeight - Image1.Height) / 2
End Sub

Private Sub Text1_Change()
End Sub

Public Sub OpenImage(imagePath As String)
    Me.Caption = imagePath
    Image1.Stretch = False
    Image1.Picture = LoadPicture(imagePath)
    Image1.Stretch = True
    SetImageInCenter
End Sub

