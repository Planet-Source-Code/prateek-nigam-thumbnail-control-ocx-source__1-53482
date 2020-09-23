VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6765
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin Project1.UserControl1 UserControl11 
      Height          =   5385
      Left            =   120
      TabIndex        =   6
      Top             =   150
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   9499
   End
   Begin VB.Label lblMove 
      AutoSize        =   -1  'True
      Caption         =   "Event"
      Height          =   195
      Left            =   1050
      TabIndex        =   5
      Top             =   6420
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Move"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   6420
      Width           =   405
   End
   Begin VB.Label lblDblClick 
      AutoSize        =   -1  'True
      Caption         =   "Event"
      Height          =   195
      Left            =   1050
      TabIndex        =   3
      Top             =   6060
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Dbl Click"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   6060
      Width           =   630
   End
   Begin VB.Label lblClick 
      AutoSize        =   -1  'True
      Caption         =   "Event"
      Height          =   195
      Left            =   1050
      TabIndex        =   1
      Top             =   5700
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Click"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   5700
      Width           =   345
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    UserControl11.SetLocation App.path
End Sub

Private Sub Form_Resize()
    UserControl11.Move UserControl11.Left, UserControl11.Top, ScaleWidth - (UserControl11.Left * 2), UserControl11.Height
End Sub

Private Sub UserControl11_click(path As String)
    lblClick = path
End Sub

Private Sub UserControl11_dblclick(path As String)
    lblDblClick.Caption = path
End Sub

Private Sub UserControl11_mouseover(path As String)
    lblMove.Caption = path
End Sub
