VERSION 5.00
Begin VB.Form frmConnect 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3345
   Icon            =   "Connect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.Frame Frame2 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2535
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2775
         Begin VB.CommandButton proceed 
            Caption         =   "Connect"
            Height          =   375
            Left            =   1800
            TabIndex        =   8
            Top             =   2040
            Width           =   855
         End
         Begin VB.CommandButton clear 
            Caption         =   "clear"
            Height          =   375
            Left            =   1080
            TabIndex        =   7
            Top             =   2040
            Width           =   615
         End
         Begin VB.FileListBox File1 
            Appearance      =   0  'Flat
            BackColor       =   &H00404000&
            ForeColor       =   &H008080FF&
            Height          =   225
            Left            =   1080
            Pattern         =   "*.ini*"
            TabIndex        =   6
            Top             =   1680
            Width           =   1605
         End
         Begin VB.TextBox passwordfield 
            Appearance      =   0  'Flat
            BackColor       =   &H00404000&
            ForeColor       =   &H008080FF&
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1080
            PasswordChar    =   "*"
            TabIndex        =   5
            Top             =   600
            Width           =   1605
         End
         Begin VB.TextBox namefield 
            Appearance      =   0  'Flat
            BackColor       =   &H00404000&
            ForeColor       =   &H008080FF&
            Height          =   285
            Left            =   1080
            TabIndex        =   4
            Top             =   240
            Width           =   1605
         End
         Begin VB.TextBox descfield 
            Appearance      =   0  'Flat
            BackColor       =   &H00404000&
            ForeColor       =   &H008080FF&
            Height          =   255
            Left            =   1080
            MultiLine       =   -1  'True
            TabIndex        =   3
            Top             =   960
            Width           =   1605
         End
         Begin VB.TextBox colorfield 
            Appearance      =   0  'Flat
            BackColor       =   &H00404000&
            ForeColor       =   &H008080FF&
            Height          =   285
            Left            =   1080
            TabIndex        =   2
            Top             =   1320
            Width           =   1605
         End
         Begin VB.Label Label1 
            BackColor       =   &H00404040&
            Caption         =   "Load Info"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Left            =   0
            TabIndex        =   13
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Left            =   0
            TabIndex        =   12
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Left            =   0
            TabIndex        =   11
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Desc"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Left            =   0
            TabIndex        =   10
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "ColorCodes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Left            =   0
            TabIndex        =   9
            Top             =   1320
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub clear_Click()
  namefield = ""
  descfield = ""
  colorfield = ""
  passwordfield = ""
End Sub

Private Sub File1_Click()
 Dim usrname As String
  Dim usrpass As String
  Dim usrcolor As String
  Dim usrdesc As String
  thefile = File1.FileName
  FileNum = FreeFile() 'Finds a freefile where it can write To
  Open App.Path & "/" & thefile For Input As FileNum 'opens the file To (input = Get data)
    Input #FileNum, usrname, usrpass, usrcolor, usrdesc 'Get data by putting Input Then the FileNumber you opened in (we used a variable FileNum) then a comma then the variable you want To store.
  Close #FileNum 'Close the FileNumber you opened...'Close' by itself will close ALL of your open files.
  namefield.Text = usrname 'sets the textbox's text = To what was is the file
  passwordfield.Text = usrpass 'sets the textbox's text = To what was is the file
  colorfield.Text = usrcolor 'sets the textbox's text = To what was is the file
  descfield.Text = usrdesc 'sets the textbox's text = To what was is the file
End Sub

Private Sub proceed_Click()
Load frmBot
frmBot.Show
frmConnect.Hide
If namefield <> "" And passwordfield <> "" And colorfield <> "" And descfield <> "" Then
On Error GoTo conerr
frmBot.sckFurc.RemoteHost = "66.28.224.193"
frmBot.sckFurc.RemotePort = "6000"
frmBot.sckFurc.Connect
End If
Exit Sub
conerr:
    box = MsgBox("Please Wait or press disconnect.", vbOKOnly, "Connect Error")
    Resume stoptrying
stoptrying:

End Sub
