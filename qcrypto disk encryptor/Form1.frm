VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1860
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "close"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "open"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "make"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4560
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "where:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Sub Command1_Click()
MkDir Mid(Drive1.Drive, 1, 2) & "\a.{00022602-0000-0000-C000-000000000046}"
FileCopy "a:\temp2.txt", Mid(Drive1.Drive, 1, 2) & "\a.{00022602-0000-0000-C000-000000000046}\temp2.txt"
MsgBox "made", , "info"
End Sub

Private Sub Command2_Click()
Dim a As String
Open "a:\temp2.txt" For Binary As #1
Input #1, ver
Close #1
Open Mid(Drive1.Drive, 1, 2) & "\a.{00022602-0000-0000-C000-000000000046}\temp2.txt" For Binary As #1
Input #1, a
Close #1
If a = ver Then
Shell ("subst g: " & Mid(Drive1.Drive, 1, 2) & "\a.{00022602-0000-0000-C000-000000000046}")
Sleep 5000
Shell ("explorer.exe g:")
Else
MsgBox "bad key"
End If
End Sub

Private Sub Command3_Click()
Shell ("subst /d g:")
MsgBox "close", , "info"
End Sub

