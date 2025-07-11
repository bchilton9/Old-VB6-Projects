VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fileeditor 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000D&
   Caption         =   "File Editor"
   ClientHeight    =   7365
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8175
   Icon            =   "fileditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   8175
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5760
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fileditor.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fileditor.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fileditor.frx":0CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fileditor.frx":1138
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Copy File"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Delete File"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "New Folder"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Rename"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6600
      Top             =   840
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000D&
      Caption         =   "Info:"
      Height          =   1575
      Left            =   4200
      TabIndex        =   10
      Top             =   5640
      Width           =   3855
      Begin VB.TextBox info2 
         Height          =   1215
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   12
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000D&
      Caption         =   "Info:"
      Height          =   1575
      Left            =   120
      TabIndex        =   9
      Top             =   5640
      Width           =   3735
      Begin VB.TextBox info1 
         Height          =   1215
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   11
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   600
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Top             =   600
      Width           =   4095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000D&
      Caption         =   "Destination "
      Height          =   3855
      Left            =   4200
      TabIndex        =   1
      Top             =   1680
      Width           =   3855
      Begin VB.FileListBox File2 
         Height          =   1845
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   3615
      End
      Begin VB.DirListBox Dir2 
         Height          =   1440
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "Source"
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   3735
      Begin VB.FileListBox File1 
         Height          =   1845
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   3495
      End
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Show:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "File:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   375
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu options 
      Caption         =   "Options"
      Begin VB.Menu delete 
         Caption         =   "Delete"
         Begin VB.Menu mnudeletefoldersource 
            Caption         =   "Source Folder"
         End
         Begin VB.Menu mnudeletedestinationfolder 
            Caption         =   "Destination Folder"
         End
      End
      Begin VB.Menu showhideinfo 
         Caption         =   "Hide Info"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "fileeditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim resizeit As Boolean, showit As Boolean
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long


Private Sub Combo1_Change()
File1.Pattern = "*." & Combo1.Text
End Sub

Private Sub Combo1_Click()
File1.Pattern = "*." & Combo1.Text
End Sub

Private Sub Dir1_Change()
File1.FileName = Dir1.Path
Dir2.Path = Dir1.Path
File2.Path = Dir2.Path
End Sub

Private Sub Dir2_Change()
File2.FileName = Dir2.Path

End Sub

Private Sub File1_Click()
Dim total As String, ext As String
Dim direct1 As String, directory As String
direct1 = Len(Dir1.Path)
If direct1 <= 3 Then
directory = "C:\" & File1.FileName
Else
directory = Dir1.Path & "\" & File1.FileName
End If
info1.Text = ""
total = Len(File1.FileName)
Text1.Text = File1.FileName
info1.Text = "FileName: " & File1.FileName & vbCrLf & _
"Size: " & Format(FileLen(directory), "###,###,###") & vbCrLf & _
"Last Date Modified: " & VBA.FileDateTime(directory)
End Sub

Private Sub File2_Click()
Dim total As String, ext As String
Dim direct1 As String, directory As String
direct1 = Len(Dir2.Path)
If direct1 <= 3 Then
directory = "C:\" & File2.FileName
Else
directory = Dir2.Path & "\" & File2.FileName
End If
info2.Text = ""
total = Len(File2.FileName)
ext = Right(directory, 3)
info2.Text = "FileName: " & File2.FileName & vbCrLf & _
"Size: " & Format(FileLen(directory), "###,###,###") & vbCrLf & _
"Last Date Modified: " & VBA.FileDateTime(directory)
End Sub

Private Sub Form_Load()
With Combo1
.AddItem "Exe"
.AddItem "Mpg"
.AddItem "Mpeg"
.AddItem "Avi"
.AddItem "Mp3"
.AddItem "Scr"
.AddItem "Com"
.AddItem "Jpg"
.AddItem "Jpeg"
.AddItem "*"
.Text = "*"
End With
Dir1.Path = "C:\"
Dir2.Path = "C:\"
fileeditor.Caption = "File Editor" & " - " & Time & " - " & Date
resizeit = False
showit = True
End Sub

Private Sub Form_Resize()
If resizeit = False Then
fileeditor.Height = "8190"
Me.Width = "8295"
showit = True
showhideinfo.Caption = "Hide Info"
Else

End If

End Sub

Private Sub mnuabout_Click()
Load credits
credits.show 1
End Sub

Private Sub mnudeletefoldersource_Click()
Dim response As String
response = MsgBox("Are you sure You Want to Delete This Folder", vbExclamation + vbOKCancel)
If response = vbOK Then
RmDir (Dir1.Path)
Dir1.Refresh
Dir2.Refresh
Dir1.ListIndex = Dir1.ListIndex - 1
End If

End Sub


Private Sub mnudeletedestinationfolder_Click()
Dim response As String
response = MsgBox("Are you sure You Want to Delete This Folder", vbExclamation + vbOKCancel)
If response = vbOK Then
RmDir (Dir2.Path)
Dir2.Refresh
Dir2.ListIndex = Dir2.ListIndex - 1
End If
End Sub


Private Sub mnuexit_Click()
End
End Sub

Private Sub showhideinfo_Click()
showit = Not showit
resizeit = True
If showit = False Then
GoTo show

Else
GoTo hideit

End If


show:
For i = Me.Height To 6480 Step -20
DoEvents
fileeditor.Height = i
fileeditor.Refresh
Next i
Me.Height = 6480
resizeit = False
showhideinfo.Caption = "Show Info"
Exit Sub

hideit:
resizeit = True
For b = Me.Height To 8175 Step 20
DoEvents
fileeditor.Height = b
fileeditor.Refresh
Next b
Me.Height = 8175
resizeit = False

showhideinfo.Caption = "Hide Info"
Exit Sub


End Sub

Private Sub Timer1_Timer()
fileeditor.Caption = ""
fileeditor.Caption = "File Editor" & " - " & Time & " - " & Date
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
If Text1.Text = "" Then
MsgBox "Select a File To Copy", vbInformation, "Select"
Else
Call copyfile
End If
Case 2
Call deletefile
Case 3
Load newfolder
newfolder.show 1
Case 4
If Text1.Text = "" Then
MsgBox "Select a File To Copy", vbInformation, "Select"
Else
Call Rename
End If

End Select
End Sub


Private Sub copyfile()
Dim direct1 As String, directory As String
Dim direct2 As String, directory2 As String
direct1 = Len(Dir1.Path)
If direct1 <= 3 Then
directory = "C:\" & File1.FileName
Else
directory = Dir1.Path & "\" & File1.FileName
End If
direct2 = Len(Dir2.Path)
If direct2 <= 3 Then
directory2 = "C:\"
Else
directory2 = Dir2.Path & "\"
End If

FileCopy directory, directory2 & Text1.Text
File1.Refresh
File2.Refresh
Beep 100, 100
DoEvents
Beep 100, 100
End Sub


Private Sub deletefile()
Dim response As String
Dim direct1 As String, directory As String
Dim direct2 As String, directory2 As String
If Text1.Text = "" Then
MsgBox "Select a File to Delete"
Exit Sub
End If
direct1 = Len(Dir1.Path)
If direct1 <= 3 Then
directory = "C:\" & File1.FileName
Else
directory = Dir1.Path & "\" & File1.FileName
End If

response = MsgBox("Are you Sure you want to completely delete: " & File1.FileName, vbYesNo, Confirm)
If response = vbYes Then Kill (directory)
File1.Refresh
File2.Refresh
Beep 500, 100
DoEvents
Beep 800, 100
End Sub

Private Sub Rename()
Dim direct1 As String, directory As String
Dim direct2 As String, directory2 As String
direct1 = Len(Dir1.Path)
If direct1 <= 3 Then
directory = "C:\" & File1.FileName
Else
directory = Dir1.Path & "\" & File1.FileName
End If
direct2 = Len(Dir2.Path)
If direct2 <= 3 Then
directory2 = "C:\"
Else
directory2 = Dir2.Path & "\"
End If

FileCopy directory, Dir1.Path & "\" & Text1.Text
Kill (directory)
File1.Refresh
File2.Refresh
Beep 350, 100
DoEvents
Beep 200, 100

End Sub
