VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWinsck.OCX"
Begin VB.Form frmTCP 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Connection Status"
   ClientHeight    =   945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTCP.frx":0000
   ScaleHeight     =   945
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox selectedmidi 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "interlude.mid"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Timer timer 
      Interval        =   30000
      Left            =   2400
      Top             =   1200
   End
   Begin MSWinsockLib.Winsock wskServer 
      Left            =   1800
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblConnect 
      BackColor       =   &H00000000&
      Caption         =   "Connecting..."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmTCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ontop As New clsOnTop
Dim mymapstr

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, _
ByVal X As Long, _
ByVal Y As Long, _
ByVal nWidth As Long, _
ByVal nHeight As Long, _
ByVal hSrcDC As Long, _
ByVal xSrc As Long, _
ByVal ySrc As Long, _
ByVal nSrcWidth As Long, _
ByVal nSrcHeight As Long, _
ByVal dwRop As Long) As Long

Private Sub cmdCancel_Click()
    Load frmLogin
    frmLogin.Show
    frmTCP.wskServer.Close
    Unload frmTCP
    
End Sub

Private Sub Form_Load()



  Set ontop = New clsOnTop
  ontop.MakeTopMost hWnd

wskServer.RemoteHost = serv

isConnected = False
isConnectedTryed = False

SetParce

End Sub

Private Sub timer_Timer()

If isConnected = False And isConnectedTryed = True Then
frmTCP.Hide
msg = MsgBox("Unabel to connect", vbCritical)
    Load frmLogin
    frmLogin.Show
    frmTCP.wskServer.Close
    Unload frmTCP
End If

End Sub

Private Sub wskServer_DataArrival(ByVal bytesTotal As Long)
Dim s As String
Dim Packet() As String
Dim i As Long

  wskServer.GetData s ', vbString, bytesTotal
  Packet = Split(s, pEnd)
  For i = 0 To UBound(Packet) - 1
    realtext Packet(i)
  Next i

End Sub

Sub realtext(txt As String)

Dim X, Y As Long

If frmDebugger.chkEnable.Value = 1 Then
frmDebugger.Text1.Text = frmDebugger.Text1.Text & vbCrLf & txt
End If

Dim Parce() As String
Parce = Split(txt, pChar)

If Parce(0) = "login" Then
lblConnect.Caption = "Loging in..."
wskServer.SendData "login" & pChar & frmLogin.txtUser.Text & pChar & frmLogin.txtPass.Text & pEnd
End If

If Parce(0) = "admin yes you are" Then
Load FrmAdmin
FrmAdmin.Show
End If

If Parce(0) = "adminmsg" Then
Parce(1) = Replace(Parce(1), "VBCRLF", vbCrLf)
msg = MsgBox("The Land's of Tidel Admin Message(s):" & Parce(1), vbOKOnly, "The Land's of Tidel Admin Message(s).")
wskServer.SendData "gotadminmsg" & pEnd
End If

If Parce(0) = "nouser" Then
    lblConnect.Caption = "Connecting..."
    frmLogin.Show
    wskServer.Close
    msd = MsgBox("Invalid Username/Password!", vbCritical)
    
    Unload frmTCP
End If


If Parce(0) = "logedin" Then
    myIndex = Parce(1)
    lblConnect.Caption = "Loged in."
    isConnected = True
    Load frmMain
    frmMain.Show
    Me.Hide

Call HTMLToRich("<FONT COLOR=#ff0000>Welcome to The Land's of Tidel!</FONT>")

startmidi = True

'Call drawMap(Player(myIndex).location)

    For i = 1 To MaxPlayers
    If Player(i).location = Player(myIndex).location Then
    PaintChar Player(i)
    End If
    
    Next i

    For Y = 0 To MapY
      For X = 0 To MapX
        With LocalMap.Tile(X, Y)

  StretchBlt frmMain.picBuffer.hdc, X * PicX, Y * PicY, 32, 32, frmMain.picObjectSelectMask.hdc, .objectoverX * 16, .objectoverY * 16, 16, 16, vbMergePaint
  StretchBlt frmMain.picBuffer.hdc, X * PicX, Y * PicY, 32, 32, frmMain.picObjectSelectB.hdc, .objectoverX * 16, .objectoverY * 16, 16, 16, vbSrcAnd

         
        End With
      Next X
    Next Y

frmMain.picBuffer.Refresh

End If

If Parce(0) = "logedout" Then
    Load frmLogin
    frmLogin.Show
    frmTCP.wskServer.Close
    Unload frmTCP
End If

If Parce(0) = "who" Then Call HTMLToRich("<FONT COLOR=#8080ff>Players on Map: </font><FONT COLOR=#ff0080>" & Parce(1) & "</FONT>")
If Parce(0) = "chat" Then Call HTMLToRich("<FONT COLOR=#00ffff>" & Parce(1) & " says, </FONT><FONT COLOR=#ffffff>" & Chr(34) & Parce(2) & Chr(34) & "</FONT>")
If Parce(0) = "youtell" Then Call HTMLToRich("<FONT COLOR=#918cc8>[ You tell </FONT><FONT COLOR=#ffff00>" & Parce(1) & "</FONT><FONT COLOR=#918cc8>, " & Chr(34) & "</FONT><FONT COLOR=#00ff00>" & Parce(2) & " </FONT><FONT COLOR=#918cc8>" & Chr(34) & " ]</FONT>")
If Parce(0) = "tellyou" Then Call HTMLToRich("<FONT COLOR=#918cc8>[ </FONT><FONT COLOR=#ffff00>" & Parce(1) & "</FONT><FONT COLOR=#918cc8> tells you, " & Chr(34) & "</FONT><FONT COLOR=#00ff00>" & Parce(2) & " </FONT><FONT COLOR=#918cc8>" & Chr(34) & " ]</FONT>")
If Parce(0) = "system" Then Call HTMLToRich("<FONT COLOR=#ff0000>" & Parce(1) & "</FONT>")

If Parce(0) = "sign" Then
frmMain.txtSign.Text = Parce(1)
frmMain.frmSign.Visible = True
End If

If Parce(0) = "inventory" Then
Dim k, l, pic, m As Integer
frmMain.txtGold.Text = Parce(89)

For k = 1 To 44
Player(myIndex).inv.invitem(k) = Parce(k)
Next k

For l = 1 To 44
pic = Parce(l + 44)
frmMain.picEquip(l).Picture = tiles.picItem(pic).Picture
Next l


frmMain.frmInventory.Visible = True
frmMain.frmBlank.Visible = True
End If


If Parce(0) = "store" Then
frmMain.frmStore.Visible = True
'frmMain.frmInventory.Visible = True
End If

If Parce(0) = "bank" Then
frmMain.frmBank.Visible = True
'frmMain.frmInventory.Visible = True
End If



If Parce(0) = "player" Then
With Player(Parce(1))
    .mask = Parce(6)
    .maskstep = 1
    .X = Val(Parce(4))
    .Y = Val(Parce(5))
    .Container = Parce(6) + 7
    .location = Parce(3)
    .Height = Parce(7)
    .Name = Parce(2)
    .offx = Parce(8)
    .offy = Parce(9)
    .Set = Parce(10)
End With
End If

If Parce(0) = "close" Then


With Player(Parce(1))
    .mask = 0
    .maskstep = 0
    .X = 0
    .Y = 0
    .Container = 0
    .location = 1
    .Height = 0
    .Name = 0
End With

Call drawMap(Player(myIndex).location)

    For i = 1 To MaxPlayers
    If Player(i).location = Player(myIndex).location Then
    PaintChar Player(i)
    End If
    Next i

    For Y = 0 To MapY
      For X = 0 To MapX
        With LocalMap.Tile(X, Y)

  StretchBlt frmMain.picBuffer.hdc, X * PicX, Y * PicY, 32, 32, frmMain.picObjectSelectMask.hdc, .objectoverX * 16, .objectoverY * 16, 16, 16, vbMergePaint
  StretchBlt frmMain.picBuffer.hdc, X * PicX, Y * PicY, 32, 32, frmMain.picObjectSelectB.hdc, .objectoverX * 16, .objectoverY * 16, 16, 16, vbSrcAnd

         
        End With
      Next X
    Next Y

frmMain.picBuffer.Refresh

End If

If Parce(0) = "move" Then

If Parce(1) = myIndex And Parce(5) <> Player(myIndex).location Then
startmidi = True
End If

With Player(Parce(1))
    .Container = Parce(2)
    .X = Val(Parce(3))
    .Y = Val(Parce(4))
    .location = Parce(5)
    .maskstep = Parce(6)
    .Height = Parce(7)
    .offx = Parce(8)
    .offy = Parce(9)
    .Set = Parce(10)
End With

Call drawMap(Player(myIndex).location)

    For i = 1 To MaxPlayers
    If Player(i).location = Player(myIndex).location Then
    PaintChar Player(i)
    End If
    
    Next i

    For Y = 0 To MapY
      For X = 0 To MapX
        With LocalMap.Tile(X, Y)

  StretchBlt frmMain.picBuffer.hdc, X * PicX, Y * PicY, 32, 32, frmMain.picObjectSelectMask.hdc, .objectoverX * 16, .objectoverY * 16, 16, 16, vbMergePaint
  StretchBlt frmMain.picBuffer.hdc, X * PicX, Y * PicY, 32, 32, frmMain.picObjectSelectB.hdc, .objectoverX * 16, .objectoverY * 16, 16, 16, vbSrcAnd

         
        End With
      Next X
    Next Y

frmMain.picBuffer.Refresh
End If

If Parce(0) = "sendinfo" Then
wskServer.SendData "newuser" & pChar & frmNew.txtUser.Text & pChar & frmNew.txtPass.Text & pChar & frmNew.txtEmail.Text & pChar & frmNew.txtMask.Text & pChar & frmNew.txtOffx.Text & pChar & frmNew.txtOffy.Text & pChar & frmNew.txtSet.Text & pEnd
End If

If Parce(0) = "usercreated" Then
'msg = MsgBox("User Created!", vbOKOnly)

Open (App.Path & "\player.dat") For Append As #1
Write #1, frmNew.txtUser.Text, frmNew.txtPass.Text
Close #1
frmLogin.lstPlayer.AddItem frmNew.txtUser.Text

    frmLogin.Show
    Unload frmTCP
    Unload frmNew
End If

End Sub

Private Sub wskServer_Close()

frmMain.txtChat.Text = frmMain.txtChat.Text & vbCrLf & "Connection Lost..."

    frmTCP.Hide
    Unload frmMain
    frmTCP.wskServer.Close
    Load frmLogin
    frmLogin.Show

End Sub
