VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl UserControl1 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   5190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6735
   ScaleHeight     =   5190
   ScaleWidth      =   6735
   Begin MSComDlg.CommonDialog cdl1 
      Left            =   120
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   5175
      Left            =   0
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim imageFilter As String
Public Sub download(fileUrl As String)
    If Ambient.UserMode Then
        On Error Resume Next
        On Error GoTo errH
        AsyncRead fileUrl, vbAsyncTypePicture
    End If
    Exit Sub
errH:
    
End Sub

Public Sub saveDownloadedPic()
    Dim ret
    On Error GoTo errH
    cdl1.FileName = ""
    cdl1.Filter = "JPEG Image (*.jpg)|*.jpg|GIF Image (*.gif)|*.gif|Bitmap Image (*.bmp)|*.bmp|Meta Picture file (*.wmf)|*.wmf|Aldus Corporation format(*.tiff)|*.tiff|WordPerfect image format (*.WPG)|*.WPG|Paint Shop Pro format (*.PSP)|*.PSP|GEM Paint format (*IMG)|*.IMG"
    cdl1.CancelError = True
    cdl1.ShowSave
    
    If Dir(cdl1.FileName) = "" Then
        SavePicture Image1.Picture, cdl1.FileName
    Else
        ret = MsgBox("A file with the name " & cdl1.FileTitle & " allready exists." & vbCrLf & "Do you want to overwrite it?", vbYesNo + vbInformation, "File Allready Exists")
        If ret = vbYes Then
            SavePicture Image1.Picture, cdl1.FileName
        Else
            Exit Sub
        End If
    End If
    Exit Sub
errH:
    
End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
    On Error Resume Next
    Image1.Picture = AsyncProp.Value
    imageFilter = Right(AsyncProp.Status, 4)
    Image1.Refresh
    UserControl.Height = Image1.Height
    UserControl.Width = Image1.Width + 50
End Sub
