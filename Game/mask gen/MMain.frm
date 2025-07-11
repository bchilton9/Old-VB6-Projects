VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "PictureBox Mask Generator"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   ScaleHeight     =   348
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   498
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "Test W/Origin."
      Height          =   255
      Left            =   150
      TabIndex        =   23
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   375
      Left            =   5880
      TabIndex        =   22
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save Corrected Image"
      Height          =   375
      Left            =   3720
      TabIndex        =   14
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save Mask Image"
      Height          =   375
      Left            =   1680
      TabIndex        =   13
      Top             =   360
      Width           =   1935
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Slow"
      Height          =   195
      Left            =   180
      TabIndex        =   12
      Top             =   4290
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6960
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Use First Pxl"
      Height          =   195
      Left            =   180
      TabIndex        =   11
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   180
      TabIndex        =   10
      Text            =   "VBWhite"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   180
      TabIndex        =   9
      Text            =   "RGB(255,255,255)"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Use First Pxl"
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Test 
      Caption         =   "Test W/Corre."
      Height          =   255
      Left            =   150
      TabIndex        =   7
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Correct Image"
      Height          =   375
      Left            =   180
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Gen Mask"
      Height          =   375
      Left            =   180
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Open1 
      Caption         =   "Open"
      Default         =   -1  'True
      Height          =   375
      Left            =   180
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.PictureBox OrPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   3735
      Left            =   1800
      ScaleHeight     =   245
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   357
      TabIndex        =   0
      Top             =   1200
      Width           =   5415
   End
   Begin VB.PictureBox CorPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   3735
      Left            =   1800
      ScaleHeight     =   245
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   357
      TabIndex        =   2
      Top             =   1200
      Width           =   5415
   End
   Begin VB.PictureBox TestPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000080FF&
      Height          =   3735
      Left            =   1800
      Picture         =   "MMain.frx":0000
      ScaleHeight     =   245
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   357
      TabIndex        =   6
      Top             =   1200
      Width           =   5415
   End
   Begin VB.PictureBox Mask 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   3735
      Left            =   1800
      ScaleHeight     =   245
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   357
      TabIndex        =   1
      Top             =   1200
      Width           =   5415
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4215
      Left            =   1680
      TabIndex        =   15
      Top             =   840
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   7435
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Original Picture"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Generated Mask"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Corrected Image"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Test Picture"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mask"
      ClipControls    =   0   'False
      Height          =   1695
      Left            =   45
      TabIndex        =   18
      Top             =   840
      Width           =   1530
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mask Color:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   960
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Image"
      Height          =   1695
      Left            =   45
      TabIndex        =   20
      Top             =   2520
      Width           =   1530
      Begin VB.Label Label2 
         Caption         =   "Correct Color:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Label Label4 
      Caption         =   "By Hou Xiong"
      Height          =   255
      Left            =   3240
      TabIndex        =   17
      Top             =   45
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "PictureBox Mask Generator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   45
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**PictureBox Mask Generator**'
'**Hou Xiong**'
'**xiong_hou@hotmail.com**'
'**houztek.cjb.net**'

'Used to check if canceled during generation
Private Declare Function GetAsyncKeyState Lib _
    "user32" (ByVal vKey As Long) As Integer

Dim Keyd As Long

Private Sub Check1_Click()
    If Check1.Value = vbChecked Then
        Text1.Enabled = False
    Else
        Text1.Enabled = True
    End If
    
    Command1.Default = True
    
End Sub

Private Sub Check2_Click()
        If Check2.Value = vbChecked Then
        Text2.Enabled = False
    Else
        Text2.Enabled = True
    End If
    
    Command2.Default = True
    
End Sub

Private Sub Command1_Click() 'generate mask
    
    Form1.MousePointer = 11
    
    Dim X As Long, Y As Long 'check for mask color
    Dim color As Long, R As Long, G As Long, B As Long
    If Check1.Value = vbChecked Then
        color = OrPic.Point(0, 0)
    Else
        If InStr(UCase(Text1.Text), "RGB") > 0 Then
            Dim starts As Long, string1 As String
            string1 = Text1.Text
            starts = InStr(string1, "(")
            string1 = Right(string1, Len(string1) - starts)
            Debug.Print Mid(string1, 1, starts - 1)
            R = Mid(string1, 1, starts - 1)
            string1 = Right(string1, Len(string1) - starts)
            starts = InStr(string1, ",")
            G = Mid(string1, 1, starts - 1)
            string1 = Right(string1, Len(string1) - starts)
            starts = InStr(string1, ")")
            B = Mid(string1, 1, starts - 1)
            color = RGB(R, G, B)
        ElseIf InStr(UCase(Text1.Text), "VB") > 0 Then
            Select Case UCase(Text1.Text)
                Case "VBWHITE"
                    color = vbWhite
                Case "VBBLACK"
                    color = vbBlack
                Case "VBBLUE"
                    color = vbBlue
                Case "VBRED"
                    color = vbRed
                Case "VBCYAN"
                    color = vbCyan
                Case "VBGREEN"
                    color = vbGreen
                Case "VBMAGENTA"
                    color = vbMagenta
                Case "VBYELLOW"
                    color = vbYellow
                Case Else
                    color = vbWhite
            End Select
        Else
            color = Text1.Text
        End If
    End If
    
    Mask.Cls 'clear mask picturebox
    
        OrPic.Visible = False
        Mask.Visible = True
        CorPic.Visible = False
        TestPic.Visible = False
        
    For Y = 0 To OrPic.ScaleHeight
        For X = 0 To OrPic.ScaleWidth
                'If point is not white, then draw black point
                If Not OrPic.Point(X, Y) = color Then Mask.PSet (X, Y), vbBlack
        'Slow on or off
        If Check3.Value = Checked Then DoEvents
        If GetAsyncKeyState(vbKeyEscape) Then
            X = OrPic.ScaleWidth
            Y = OrPic.ScaleHeight
            Form1.MousePointer = 0
        End If
        Next
    Next
    
    Form1.MousePointer = 0
End Sub

'What I mean by correct picture is that it has to have a white background
'otherwise the blt'ed image won't look normal
Private Sub Command2_Click() 'Correct picture
    Form1.MousePointer = 11
    
    Dim X As Long, Y As Long
    Dim color As Long, R As Long, G As Long, B As Long
    If Check2.Value = vbChecked Then
        color = OrPic.Point(0, 0)
    Else
        If InStr(UCase(Text1.Text), "RGB") > 0 Then
            Dim starts As Long, string1 As String
            string1 = Text1.Text
            starts = InStr(string1, "(")
            string1 = Right(string1, Len(string1) - starts)
            Debug.Print Mid(string1, 1, starts - 1)
            R = Mid(string1, 1, starts - 1)
            string1 = Right(string1, Len(string1) - starts)
            starts = InStr(string1, ",")
            G = Mid(string1, 1, starts - 1)
            string1 = Right(string1, Len(string1) - starts)
            starts = InStr(string1, ")")
            B = Mid(string1, 1, starts - 1)
            color = RGB(R, G, B)
        ElseIf InStr(UCase(Text1.Text), "VB") > 0 Then
            Select Case UCase(Text1.Text)
                Case "VBWHITE"
                    color = vbWhite
                Case "VBBLACK"
                    color = vbBlack
                Case "VBBLUE"
                    color = vbBlue
                Case "VBRED"
                    color = vbRed
                Case "VBCYAN"
                    color = vbCyan
                Case "VBGREEN"
                    color = vbGreen
                Case "VBMAGENTA"
                    color = vbMagenta
                Case "VBYELLOW"
                    color = vbYellow
                Case Else
                    color = vbWhite
            End Select
        Else
            color = Text1.Text
        End If
    End If
    
    CorPic.Cls
    
        OrPic.Visible = False
        Mask.Visible = False
        CorPic.Visible = True
        TestPic.Visible = False
    
    For Y = 0 To OrPic.ScaleHeight
        For X = 0 To OrPic.ScaleWidth
                'if color of point is = to mask color then draw white point
                If OrPic.Point(X, Y) = color Then
                CorPic.PSet (X, Y), vbWhite
                Else
                'otherwise keep color as is
                CorPic.PSet (X, Y), OrPic.Point(X, Y)
                End If
        If Check3.Value = Checked Then DoEvents
        If GetAsyncKeyState(vbKeyEscape) Then
            X = OrPic.ScaleWidth
            Y = OrPic.ScaleHeight
            Form1.MousePointer = 0
        End If
        Next
    Next
    
    Form1.MousePointer = 0

End Sub



Private Sub Command3_Click()
    CommonDialog1.Filter = "Pictures|*.bmp;*.dib;*.jpp;*.gif;*.ico;*.cur;*.wmf;*.emf|All Files|*.*"
    CommonDialog1.DefaultExt = "*.bmp"
    CommonDialog1.ShowSave
    If CommonDialog1.FileName <> "" Then
        SavePicture Mask.Image, CommonDialog1.FileName
    End If
End Sub

Private Sub Command4_Click() 'Save picture
    CommonDialog1.Filter = "Pictures|*.bmp;*.dib;*.jpp;*.gif;*.ico;*.cur;*.wmf;*.emf|All Files|*.*"
    CommonDialog1.DefaultExt = "*.bmp"
    CommonDialog1.ShowSave
    If CommonDialog1.FileName <> "" Then
        SavePicture CorPic.Image, CommonDialog1.FileName
    End If
End Sub

Private Sub Command5_Click()
    Unload Me
End Sub

Private Sub Command6_Click() 'test pic with original picture
    TestPic.Cls
    TestPic.PaintPicture Mask.Image, 0, 0, OrPic.ScaleWidth, OrPic.ScaleHeight, 0, 0, OrPic.ScaleWidth, OrPic.ScaleHeight, vbMergePaint
    TestPic.PaintPicture OrPic.Image, 0, 0, OrPic.ScaleWidth, OrPic.ScaleHeight, 0, 0, OrPic.ScaleWidth, OrPic.ScaleHeight, vbSrcAnd
    TestPic.Refresh
        OrPic.Visible = False
        Mask.Visible = False
        CorPic.Visible = False
        TestPic.Visible = True
End Sub

Private Sub Form_Load()
    Randomize
    Check1.Value = vbChecked
    Check2.Value = vbChecked
    Text1.Enabled = False
    Text2.Enabled = False
    TestPic.BackColor = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
    Mask.Visible = False
    CorPic.Visible = False
    TestPic.Visible = False
    Open1.Default = True
    
End Sub

Private Sub Form_Terminate()
    Form1.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Confirm exit
    If MsgBox("Are you sure you want to exit?" & vbCr & "All modifications, if any, will be lost.", vbQuestion + vbYesNo, "Exit?") = vbYes Then
    Form1.MousePointer = 0
    End
    Else
    Cancel = 1
    End If
End Sub

Private Sub Open1_Click() 'open file
    CommonDialog1.Filter = "Pictures|*.bmp;*.dib;*.jpp;*.gif;*.ico;*.cur;*.wmf;*.emf|All Files|*.*"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        OrPic.Picture = LoadPicture(CommonDialog1.FileName)
        Mask.Width = OrPic.Width
        Mask.Height = OrPic.Height
        CorPic.Width = OrPic.Width
        CorPic.Height = OrPic.Height
        TestPic.Width = OrPic.Width
        TestPic.Height = OrPic.Height
        TabStrip1.Width = OrPic.Width + 15
        TabStrip1.Height = OrPic.Height + 31
    End If
        OrPic.Visible = True
        Mask.Visible = False
        CorPic.Visible = False
        TestPic.Visible = False
End Sub

Private Sub TabStrip1_Click() 'Change picturebox view
    If TabStrip1.SelectedItem = "&Original Picture" Then
        OrPic.Visible = True
        Mask.Visible = False
        CorPic.Visible = False
        TestPic.Visible = False
    ElseIf TabStrip1.SelectedItem = "&Generated Mask" Then
        OrPic.Visible = False
        Mask.Visible = True
        CorPic.Visible = False
        TestPic.Visible = False
    ElseIf TabStrip1.SelectedItem = "&Corrected Image" Then
        OrPic.Visible = False
        Mask.Visible = False
        CorPic.Visible = True
        TestPic.Visible = False
    ElseIf TabStrip1.SelectedItem = "&Test Picture" Then
        OrPic.Visible = False
        Mask.Visible = False
        CorPic.Visible = False
        TestPic.Visible = True
    End If
End Sub

Private Sub Test_Click() 'test pic with corrected image
    TestPic.Cls
    TestPic.PaintPicture Mask.Image, 0, 0, OrPic.ScaleWidth, OrPic.ScaleHeight, 0, 0, OrPic.ScaleWidth, OrPic.ScaleHeight, vbMergePaint
    TestPic.PaintPicture CorPic.Image, 0, 0, OrPic.ScaleWidth, OrPic.ScaleHeight, 0, 0, OrPic.ScaleWidth, OrPic.ScaleHeight, vbSrcAnd
    TestPic.Refresh
        OrPic.Visible = False
        Mask.Visible = False
        CorPic.Visible = False
        TestPic.Visible = True
End Sub

Private Sub Text1_GotFocus()
    Command1.Default = True
End Sub

Private Sub Text2_GotFocus()
    Command2.Default = True
End Sub

'Please give me credits where possible
