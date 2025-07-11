VERSION 5.00
Begin VB.Form frmGen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patch File Creator"
   ClientHeight    =   4680
   ClientLeft      =   4230
   ClientTop       =   2625
   ClientWidth     =   6345
   Icon            =   "frmGen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6345
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optType 
      Caption         =   "Midis"
      Height          =   255
      Index           =   4
      Left            =   1680
      TabIndex        =   10
      Top             =   4080
      Width           =   1455
   End
   Begin VB.OptionButton optType 
      Caption         =   "Skin"
      Height          =   255
      Index           =   3
      Left            =   1680
      TabIndex        =   9
      Top             =   3840
      Width           =   1455
   End
   Begin VB.OptionButton optType 
      Caption         =   "Patcher"
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   8
      Top             =   4320
      Width           =   1455
   End
   Begin VB.OptionButton optType 
      Caption         =   "Client"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   7
      Top             =   4080
      Width           =   1455
   End
   Begin VB.OptionButton optType 
      Caption         =   "Maps"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   6
      Top             =   3840
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Create Patch File"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   1320
      Width           =   375
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   3255
   End
   Begin VB.FileListBox File 
      Height          =   2040
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   2295
   End
   Begin VB.DirListBox Dir 
      Height          =   1440
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private fso As New FileSystemObject

Private Sub Command1_Click()
List1.AddItem File.Path & "\" & File.FileName
End Sub

Private Sub Command2_Click()
If List1.ListIndex <> -1 Then
List1.RemoveItem (List1.ListIndex)
End If
End Sub

Private Sub Command3_Click()
'List1.ListCount List1.Text
Dim i As Integer

If optType(1).Value = True Then
saveto = App.Path & "\patch\client.pch"
savetoback = App.Path & "\client.bak"
ElseIf optType(2).Value = True Then
saveto = App.Path & "\patch\patcher.pch"
savetoback = App.Path & "\patcher.bak"
ElseIf optType(3).Value = True Then
saveto = App.Path & "\patch\skin.pch"
savetoback = App.Path & "\skin.bak"
ElseIf optType(4).Value = True Then
saveto = App.Path & "\patch\midi.pch"
savetoback = App.Path & "\midi.bak"
Else
saveto = App.Path & "\patch\maps.pch"
savetoback = App.Path & "\maps.bak"
End If

Open saveto For Output As #4
Print #4, ""
Close #4
Open savetoback For Output As #3
Print #3, ""
Close #3

Open saveto For Output As #1
Open savetoback For Output As #2

For i = 0 To List1.ListCount - 1

GetFile = List1.List(i)

isfile = fso.FileExists(GetFile)
FileName = fso.GetFileName(GetFile)
LenFile = FileLen(GetFile)
FileVer = fso.GetFileVersion(GetFile)
lastMod = FileDateTime(GetFile)

Print #1, Chr(34) & FileName & Chr(34) & "," & Chr(34) & LenFile & Chr(34) & "," & Chr(34) & FileVer & Chr(34) & "," & Chr(34) & lastMod & Chr(34)
Print #2, Chr(34) & List1.List(i) & Chr(34)

Next i

Close #1, #2

msg = MsgBox("Patch file saved as " & Chr(34) & saveto & Chr(34) & ".", vbOKOnly, "Patch File Saved")

End Sub

Private Sub Dir_Change()
File.Path = Dir.Path
End Sub

Private Sub Form_Load()
Dir.Path = App.Path

optType_Click (0)

End Sub

Private Sub optType_Click(Index As Integer)

List1.Clear

If optType(1).Value = True Then

savetobackopen = App.Path & "\client.bak"
ElseIf optType(2).Value = True Then

savetobackopen = App.Path & "\patcher.bak"
ElseIf optType(3).Value = True Then

savetobackopen = App.Path & "\skin.bak"
ElseIf optType(4).Value = True Then

savetobackopen = App.Path & "\midi.bak"
Else

savetobackopen = App.Path & "\maps.bak"
End If

On Error GoTo skip
Open savetobackopen For Input As #1
Do Until (EOF(1))
   Input #1, BakFileName
   List1.AddItem BakFileName
Loop
Close #1
skip:

End Sub
