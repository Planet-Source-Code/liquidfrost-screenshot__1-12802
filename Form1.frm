VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Take Screenshot"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   6855
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   4680
      Width           =   7095
   End
   Begin VB.Frame Frame2 
      Height          =   4575
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7095
      Begin VB.Image Image1 
         Height          =   4215
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   6855
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Sub Command1_Click()
Form1.Hide
Screenshot "c:\windows\desktop\Screenshot.bmp"
Image1.Picture = LoadPicture("c:\windows\desktop\Screenshot.bmp")
Form1.Show
End Sub
Public Function Screenshot(ByVal Destination$) As Boolean

On Error Resume Next
DoEvents
Call keybd_event(vbKeySnapshot, 1, 0, 0)
DoEvents
SavePicture Clipboard.GetData(vbCFBitmap), Destination$
Screenshot = True
End Function


Private Sub Form_Unload(Cancel As Integer)
End
End Sub

