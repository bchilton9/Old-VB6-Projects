VERSION 5.00
Begin VB.Form doAdd 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   555
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   555
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtO2 
      Height          =   285
      Left            =   1560
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtO1 
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtY 
      Height          =   285
      Left            =   2400
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtX 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox cmbFunc 
      Height          =   315
      ItemData        =   "doAdd.frx":0000
      Left            =   120
      List            =   "doAdd.frx":0013
      TabIndex        =   0
      Text            =   "Please select one..."
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label lblF1 
      Alignment       =   1  'Right Justify
      Caption         =   "floor 1:"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblF2 
      Alignment       =   1  'Right Justify
      Caption         =   "floor 2:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblO2 
      Alignment       =   1  'Right Justify
      Caption         =   "Object 2:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblO1 
      Alignment       =   1  'Right Justify
      Caption         =   "Object 1:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblY 
      Alignment       =   2  'Center
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
      Left            =   2160
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
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
      Left            =   1440
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblPos 
      Alignment       =   1  'Right Justify
      Caption         =   "Position:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "doAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ntxtO1, ntxtO2, ntxtX, ntxtY As Boolean

Private Sub cmbFunc_click()
doAdd.Height = "2610"
If cmbFunc.Text = "Place object # at X,Y." Then
        txtY.Visible = True
        txtX.Visible = True
        lblY.Visible = True
        lblX.Visible = True
        lblPos.Visible = True
        lblO1.Visible = True
        txtO1.Visible = True
        lblO2.Visible = False
        txtO2.Visible = False
        lblF1.Visible = False
        lblF2.Visible = False
        ntxtO1 = True
        ntxtO2 = False
        ntxtX = True
        ntxtY = True
ElseIf cmbFunc.Text = "Swap object's # and # at X,Y." Then
        txtY.Visible = True
        txtX.Visible = True
        lblY.Visible = True
        lblX.Visible = True
        lblPos.Visible = True
        lblO1.Visible = True
        txtO1.Visible = True
        lblO2.Visible = True
        txtO2.Visible = True
        lblF1.Visible = False
        lblF2.Visible = False
        ntxtO1 = True
        ntxtO2 = True
        ntxtX = True
        ntxtY = True
ElseIf cmbFunc.Text = "Swap object's # and #." Then
        txtY.Visible = False
        txtX.Visible = False
        lblY.Visible = False
        lblX.Visible = False
        lblPos.Visible = False
        lblO1.Visible = True
        txtO1.Visible = True
        lblO2.Visible = True
        txtO2.Visible = True
        lblF1.Visible = False
        lblF2.Visible = False
        ntxtO1 = True
        ntxtO2 = True
        ntxtX = False
        ntxtY = False
ElseIf cmbFunc.Text = "Swap floor's # and # at X,Y." Then
        txtY.Visible = True
        txtX.Visible = True
        lblY.Visible = True
        lblX.Visible = True
        lblPos.Visible = True
        lblO1.Visible = False
        txtO1.Visible = True
        lblO2.Visible = False
        txtO2.Visible = True
        lblF1.Visible = True
        lblF2.Visible = True
        ntxtO1 = True
        ntxtO2 = True
        ntxtX = True
        ntxtY = True
ElseIf cmbFunc.Text = "Swap floor's # and #." Then
        txtY.Visible = False
        txtX.Visible = False
        lblY.Visible = False
        lblX.Visible = False
        lblPos.Visible = False
        lblO1.Visible = False
        txtO1.Visible = True
        lblO2.Visible = False
        txtO2.Visible = True
        lblF1.Visible = True
        lblF2.Visible = True
        ntxtO1 = True
        ntxtO2 = True
        ntxtX = False
        ntxtY = False
End If
End Sub


Private Sub cmdSave_Click()
If ntxtO1 = True And txtO1 = "" Then
mssg = MsgBox("Please enter all the felds!", vbOKOnly)
ElseIf ntxtO2 = True And txtO2 = "" Then
mssg = MsgBox("Please enter all the felds!", vbOKOnly)
ElseIf ntxtX = True And txtX = "" Then
mssg = MsgBox("Please enter all the felds!", vbOKOnly)
ElseIf ntxtY = True And txtY = "" Then
mssg = MsgBox("Please enter all the felds!", vbOKOnly)
Else


If cmbFunc.Text = "Place object # at X,Y." Then
Open frmMain.txtTemp For Append As #1
Write #1, "    (3:2) at position (" & txtX & "," & txtY & ") on the map,"
Write #1, "    (5:4) place object type " & txtO1 & "."
Close #1
capt = "Place object " & txtO1 & " at " & txtX & "," & txtY & "."

ElseIf cmbFunc.Text = "Swap object's # and # at X,Y." Then
Open frmMain.txtTemp For Append As #1
Write #1, "    (3:2) at position (" & txtX & "," & txtY & ") on the map,"
Write #1, "    (5:6) swap object types " & txtO1 & " and " & txtO2 & "."
Close #1
capt = "Swap object's " & txtO1 & " and " & txtO2 & " at " & txtX & "," & txtY & "."

ElseIf cmbFunc.Text = "Swap object's # and #." Then
Open frmMain.txtTemp For Append As #1
Write #1, "    (5:6) swap object types " & txtO1 & " and " & txtO2 & "."
Close #1
capt = "Swap object's " & txtO1 & " and " & txtO2 & "."

ElseIf cmbFunc.Text = "Swap floor's # and # at X,Y." Then
Open frmMain.txtTemp For Append As #1
Write #1, "    (3:2) at position (" & txtX & "," & txtY & ") on the map,"
Write #1, "    (5:3) swap floor types " & txtO1 & " and " & txtO2 & "."
Close #1
capt = "Swap floor's " & txtO1 & " and " & txtO2 & " at " & txtX & "," & txtY & "."

ElseIf cmbFunc.Text = "Swap floor's # and #." Then
Open frmMain.txtTemp For Append As #1
Write #1, "    (5:3) swap floor types " & txtO1 & " and " & txtO2 & "."
Close #1
capt = "Swap floor's " & txtO1 & " and " & txtO2 & "."
End If


If frmMain.txtFrom = "switchon" Then Switch.txtOn = Switch.txtOn & capt & vbCrLf
If frmMain.txtFrom = "switchoff" Then Switch.txtOff = Switch.txtOff & capt & vbCrLf
Unload doAdd
End If
End Sub
