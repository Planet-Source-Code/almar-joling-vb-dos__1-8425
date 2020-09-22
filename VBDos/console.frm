VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   7065
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   2040
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   3000
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

