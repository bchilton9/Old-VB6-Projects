VERSION 5.00
Begin VB.Form frmAddFunc 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   1785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstFunc 
      BackColor       =   &H00FFC0C0&
      Height          =   1035
      ItemData        =   "frmAddFunc.frx":0000
      Left            =   2760
      List            =   "frmAddFunc.frx":0002
      TabIndex        =   21
      Top             =   2400
      Width           =   2655
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Delete"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Edit"
      Enabled         =   0   'False
      Height          =   255
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Add"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2760
      Width           =   735
   End
   Begin VB.ListBox lstComds 
      BackColor       =   &H00FFC0C0&
      Height          =   2205
      ItemData        =   "frmAddFunc.frx":0004
      Left            =   120
      List            =   "frmAddFunc.frx":0006
      TabIndex        =   13
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdDone 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Done"
      Height          =   255
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3120
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Timer"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   1200
      Width           =   2895
   End
   Begin VB.CheckBox chkSign 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Sign."
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   960
      Width           =   2895
   End
   Begin VB.CheckBox chkSay 
      BackColor       =   &H00FFFFC0&
      Caption         =   "When something is sead."
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   720
      Width           =   2895
   End
   Begin VB.CheckBox chkWhis 
      BackColor       =   &H00FFFFC0&
      Caption         =   "When someone whispers the bot."
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   480
      Width           =   2895
   End
   Begin VB.CommandButton cmdAddf 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Add"
      Height          =   255
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Save"
      Height          =   255
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Commds"
      Height          =   255
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox txtY 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   4560
      MaxLength       =   3
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtX 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   3720
      MaxLength       =   3
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtSend 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   3360
      TabIndex        =   2
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox txtRec 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   3360
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblY 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   20
      Top             =   1680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblX 
      BackColor       =   &H00FFFFC0&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   19
      Top             =   1680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblAdd 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   2040
      TabIndex        =   18
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Command's"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "Bot Send's:"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   1
      Top             =   2040
      Width           =   1095
   End
End
Attribute VB_Name = "frmAddFunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim isok As Boolean
Private Sub chkSay_Click()
If chkSay.Value = Checked Then
lblAdd.Caption = "Bot Recive's:"
lblX.Visible = False
lblY.Visible = False
txtX.Visible = False
txtY.Visible = False
txtRec.Visible = True
chkWhis.Value = 0
chkSign.Value = 0
End If
End Sub

Private Sub chkSign_Click()
If chkSign.Value = Checked Then
lblAdd.Caption = "Sign Location:"
lblX.Visible = True
lblY.Visible = True
txtX.Visible = True
txtY.Visible = True
txtRec.Visible = False
chkSay.Value = 0
chkWhis.Value = 0
End If
End Sub

Private Sub chkWhis_Click()
If chkWhis.Value = Checked Then
lblAdd.Caption = "Bot Recive's:"
lblX.Visible = False
lblY.Visible = False
txtX.Visible = False
txtY.Visible = False
txtRec.Visible = True
chkSay.Value = 0
chkSign.Value = 0
End If
End Sub

Private Sub cmdAdd_Click()
frmAddFunc.Width = 5595
End Sub

Private Sub cmdDelete_Click()
On Error GoTo error
del = Right(lstComds.Text, Len(lstComds.Text) - 7)
Open frmBot.bload & ".lst" For Input As #1
Open "temp.lst" For Output As #2
Do Until EOF(1)
Input #1, typ, x, y, lin, cmmd
If lin <> del Then
Write #2, typ, x, y, lin, cmmd
End If
Loop
Close #1, #2

Open frmBot.bload & ".lst" For Output As #1
Open "temp.lst" For Input As #2
Do Until EOF(2)
Input #2, typ, x, y, lin, cmmd
If lin <> del Then
Write #1, typ, x, y, lin, cmmd
End If
Loop
Close #1, #2

Open "temp.lst" For Output As #2
Write #2, ""
Close #2

lstComds.Clear
Open frmBot.bload & ".lst" For Input As #1
Do Until EOF(1)
Input #1, typ, x, y, lin, cmmd
If typ = "wh" Then
    lstComds.AddItem "Whis - " & lin
Else
If typ = "si" Then
    lstComds.AddItem "Sign - " & lin
Else
If typ = "sa" Then
    lstComds.AddItem "Say  - " & lin
End If
End If
End If
Loop
Close #1
Exit Sub
error:
    Close #1, #2
    Resume stoptrying
stoptrying:
End Sub

Private Sub cmdDone_Click()
Unload frmAddFunc
End Sub

Private Sub cmdAddf_Click()
On Error GoTo error
If txtSend <> "" Then
Open "temp.lst" For Append As #1
Write #1, txtSend
Close #1
lstFunc.AddItem txtSend
txtSend = ""
isok = True
End If
Exit Sub
error:
    Close #1, #2
    Resume stoptrying
stoptrying:
End Sub

Private Sub cmdSave_Click()
On Error GoTo error
If chkSign.Value = 1 And txtX = "" Then
        box = MsgBox("Please enter the sign location!", vbOKOnly)
Else
If chkSign.Value = 1 And txtY = "" Then
        box = MsgBox("Please enter the sign location!", vbOKOnly)
Else
If chkSign.Value = 0 And txtRec = "" Then
    box = MsgBox("Please enter what the bot recives!", vbOKOnly)
Else
If isok = False Then
    box = MsgBox("Please enter a responce!", vbOKOnly)
Else

txtRec = LCase(txtRec)
Open "temp.lst" For Input As #1
Do Until EOF(1)
Input #1, lin
If lin <> "" Then
    If mess = "" Then
    mess = lin & ";"
    Else
    mess = mess & lin & ";"
    End If
End If
Loop
Close #1
Open frmBot.bload & ".lst" For Append As #1
If chkWhis.Value = 1 Then
    Write #1, "wh", 0, 0, txtRec, mess
    lstComds.AddItem "Whis - " & txtRec
Else
If chkSay.Value = 1 Then
    Write #1, "sa", 0, 0, txtRec, mess
    lstComds.AddItem "Say  - " & txtRec
Else
If chkSign.Value = 1 Then
    iX = Int(txtX)
    iY = Int(txtY)
    Write #1, "si", iX, iY, "(" & txtX & "," & txtY & ")", mess
    lstComds.AddItem "Sign - " & "(" & txtX & "," & txtY & ")"
End If
End If
End If
Close #1
Open "temp.lst" For Output As #2
Write #2, ""
Close #2
lstFunc.Clear
txtSend = ""
txtX = ""
txtY = ""
txtRec = ""
isok = False
chkSign.Value = 0
chkSay.Value = 0
chkWhis.Value = 0
End If
End If
End If
End If
Exit Sub
error:
    Close #1, #2
    Resume stoptrying
stoptrying:
End Sub

Private Sub Command1_Click()
comds.Show
End Sub

Private Sub Form_Load()
On Error GoTo error
isok = False
Open frmBot.bload & ".lst" For Input As #1
Do Until EOF(1)
Input #1, typ, x, y, lin, cmmd
If typ = "wh" Then
    lstComds.AddItem "Whis - " & lin
Else
If typ = "si" Then
    lstComds.AddItem "Sign - " & lin
Else
If typ = "sa" Then
    lstComds.AddItem "Say  - " & lin
End If
End If
End If
Loop
Close #1
Exit Sub
error:
    Close #1, #2
    Resume stoptrying
stoptrying:
End Sub
