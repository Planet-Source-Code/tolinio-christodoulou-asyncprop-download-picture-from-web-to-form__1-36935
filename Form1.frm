VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   Caption         =   "Download Picture From a Website directly to the form"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8880
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   8880
   StartUpPosition =   3  'Windows Default
   Begin Project1.UserControl1 UserControl11 
      Height          =   5175
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   9128
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frame1"
      Height          =   855
      Left            =   -120
      TabIndex        =   0
      Top             =   -240
      Width           =   10815
      Begin VB.CommandButton Command2 
         Caption         =   "Save Picture"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         ItemData        =   "Form1.frx":08CA
         Left            =   2640
         List            =   "Form1.frx":08E9
         TabIndex        =   2
         Top             =   360
         Width           =   6255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Download"
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command1_Click
    End If
End Sub

Private Sub Command1_Click()
    UserControl11.download Combo1.Text
    If Me.WindowState = vbNormal Then
        If UserControl11.Width <= 8340 And UserControl11.Height <= 6420 + Frame1.Height Then
            Me.Width = 8340
            Me.Height = 6420 + Frame1.Height
        Else
            Me.Move Me.Left, Me.Top, UserControl11.Width + 100, UserControl11.Height + Frame1.Height + 450
            DoEvents
        End If
    End If
End Sub

Private Sub Command2_Click()
    UserControl11.saveDownloadedPic
End Sub

Private Sub Form_Load()
    Combo1.ListIndex = 0
End Sub

Private Sub Form_Resize()
    Frame1.Width = Me.ScaleWidth + 200
    Combo1.Width = Me.ScaleWidth - (Command1.Width + Command2.Width + 400)
End Sub
