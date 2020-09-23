VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form TestSuite 
   Caption         =   "qcrypto disk encryptor!"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   960
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "X"
      TabIndex        =   3
      Text            =   "password"
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "verifier"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Text            =   "login"
      Top             =   120
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "make key"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "password:"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "login:"
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   240
      Width           =   375
   End
End
Attribute VB_Name = "TestSuite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetVolumeLabel Lib "kernel32" Alias "SetVolumeLabelA" (ByVal lpRootPathName As String, ByVal lpVolumeName As String) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Dim md5Test As MD5
Dim a As String
Dim b As String
Dim c As String
Dim d As String
Private Sub Command1_Click()
Dim temp As String
Set md5Test = New MD5
Dim Serial As Long, VName As String, FSName As String
VName = String$(255, Chr$(0))
GetVolumeInformation "a:\", VName, 255, Serial, 0, 0, vbNullString, 0
c = Trim(Str$(Serial))
c = md5Test.DigestStrToHexStr(c)
d = enc(c, md5Test.DigestStrToHexStr(Text2.Text))
MsgBox "choose 3 file to make the randome key", , "instruction"
CommonDialog1.ShowOpen
a = md5Test.DigestFileToHexStr(CommonDialog1.FileName)
CommonDialog1.ShowOpen
a = a & md5Test.DigestFileToHexStr(CommonDialog1.FileName)
CommonDialog1.ShowOpen
a = a & md5Test.DigestFileToHexStr(CommonDialog1.FileName)
Open "a:\temp.txt" For Binary As #1
Put #1, , a
Close #1
b = md5Test.DigestFileToHexStr("a:\temp.txt")
temp = enc(d, b)
Open "a:\temp2.txt" For Binary As #1
Put #1, 1, temp
Close #1

SetVolumeLabel "a:\", Text1.Text
End Sub

Private Sub Command2_Click()
Dim temp As String
Dim ver As String


Set md5Test = New MD5
Dim Serial As Long, VName As String, FSName As String
VName = String$(255, Chr$(0))
GetVolumeInformation "a:\", VName, 255, Serial, 0, 0, vbNullString, 0
c = Trim(Str$(Serial))
c = md5Test.DigestStrToHexStr(c)
d = enc(c, md5Test.DigestStrToHexStr(Text2.Text))
b = md5Test.DigestFileToHexStr("a:\temp.txt")
temp = enc(d, b)
Open "a:\temp2.txt" For Binary As #1
Input #1, ver
Close #1



Text3.Text = VName
If Text3.Text = UCase(Text1.Text) Then
Text1.Text = ver
Text2.Text = temp
If Text1.Text = Text2.Text Then
Text1.Text = ""
Text2.Text = ""
MsgBox "accepted", , "sucessful"
Form1.Show
Else
MsgBox "bad password or disk", , "error"
End If
Else
MsgBox "bad login", , "error"
End If


End Sub

Public Function enc(s As String, t As String)


Dim temp As String
Dim i As Integer
Dim location As Integer
temp$ = ""
For i% = 1 To Len(s)
    location% = (i% Mod Len(t)) + 1
    temp$ = temp$ + Chr$(Asc(Mid$(s, i%, 1)) Xor Asc(Mid$(t, location%, 1)))
Next i%
enc = temp
End Function
