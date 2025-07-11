VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmBot 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mesto"
   ClientHeight    =   4695
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   5415
   Icon            =   "frmBot.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmBot.frx":0442
   ScaleHeight     =   4695
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Items"
      Height          =   2535
      Left            =   5640
      TabIndex        =   42
      Top             =   4200
      Width           =   2535
      Begin VB.TextBox txtItemType 
         DataField       =   "type"
         DataSource      =   "items"
         Height          =   285
         Left            =   960
         TabIndex        =   57
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox txtItemDmg 
         DataField       =   "Dmg"
         DataSource      =   "items"
         Height          =   285
         Left            =   120
         TabIndex        =   56
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtItemEquip 
         DataField       =   "Equip"
         DataSource      =   "items"
         Height          =   285
         Left            =   1920
         TabIndex        =   54
         Top             =   720
         Width           =   495
      End
      Begin VB.Data items 
         Caption         =   "Items"
         Connect         =   "Access"
         DatabaseName    =   "F:\vbprojects\Mesto\mesto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Items"
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtItemID 
         DataField       =   "ID"
         DataSource      =   "items"
         Height          =   285
         Left            =   120
         TabIndex        =   49
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtItemName 
         DataField       =   "Name"
         DataSource      =   "items"
         Height          =   285
         Left            =   120
         TabIndex        =   48
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtItemDesc 
         DataField       =   "Description"
         DataSource      =   "items"
         Height          =   285
         Left            =   120
         TabIndex        =   47
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox txtItemBuy 
         DataField       =   "buy"
         DataSource      =   "items"
         Height          =   285
         Left            =   720
         TabIndex        =   46
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtItemSell 
         DataField       =   "sell"
         DataSource      =   "items"
         Height          =   285
         Left            =   1320
         TabIndex        =   45
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtItemHP 
         DataField       =   "HP"
         DataSource      =   "items"
         Height          =   285
         Left            =   120
         TabIndex        =   44
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtItemMana 
         DataField       =   "Mana"
         DataSource      =   "items"
         Height          =   285
         Left            =   1320
         TabIndex        =   43
         Top             =   1800
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Members"
      Height          =   3975
      Left            =   5640
      TabIndex        =   19
      Top             =   120
      Width           =   2535
      Begin VB.TextBox txtDmg 
         DataField       =   "Damage"
         DataSource      =   "mbr"
         Height          =   285
         Left            =   120
         TabIndex        =   58
         Top             =   3600
         Width           =   495
      End
      Begin VB.TextBox txtWeapon 
         DataField       =   "weapon"
         DataSource      =   "mbr"
         Height          =   285
         Left            =   1320
         TabIndex        =   55
         Top             =   2880
         Width           =   495
      End
      Begin VB.TextBox txtEquip4 
         DataField       =   "Equip4"
         DataSource      =   "mbr"
         Height          =   285
         Left            =   1920
         TabIndex        =   53
         Top             =   3240
         Width           =   495
      End
      Begin VB.TextBox txtEquip3 
         DataField       =   "Equip3"
         DataSource      =   "mbr"
         Height          =   285
         Left            =   1320
         TabIndex        =   52
         Top             =   3240
         Width           =   495
      End
      Begin VB.TextBox txtEquip2 
         DataField       =   "Equip2"
         DataSource      =   "mbr"
         Height          =   285
         Left            =   720
         TabIndex        =   51
         Top             =   3240
         Width           =   495
      End
      Begin VB.TextBox txtEquip1 
         DataField       =   "Equip1"
         DataSource      =   "mbr"
         Height          =   285
         Left            =   120
         TabIndex        =   50
         Top             =   3240
         Width           =   495
      End
      Begin VB.Data mbr 
         Caption         =   "Members"
         Connect         =   "Access"
         DatabaseName    =   "F:\vbprojects\Mesto\mesto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Members"
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtID 
         DataField       =   "ID"
         DataSource      =   "mbr"
         Height          =   285
         Left            =   120
         TabIndex        =   41
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtName 
         DataField       =   "Name"
         DataSource      =   "mbr"
         Height          =   285
         Left            =   720
         TabIndex        =   40
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtlvl 
         DataField       =   "Level"
         DataSource      =   "mbr"
         Height          =   285
         Left            =   720
         TabIndex        =   39
         Top             =   2880
         Width           =   495
      End
      Begin VB.TextBox txtDate 
         DataField       =   "Date"
         DataSource      =   "mbr"
         Height          =   285
         Left            =   120
         TabIndex        =   38
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtInv1 
         DataField       =   "Inv1"
         DataSource      =   "mbr"
         Height          =   285
         Left            =   120
         TabIndex        =   37
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox txtInv1q 
         DataField       =   "Inv1q"
         DataSource      =   "mbr"
         Height          =   285
         Left            =   720
         TabIndex        =   36
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox txtInv2 
         DataField       =   "Inv2"
         DataSource      =   "mbr"
         Height          =   285
         Left            =   120
         TabIndex        =   35
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txtInv2q 
         DataField       =   "Inv2q"
         DataSource      =   "mbr"
         Height          =   285
         Left            =   720
         TabIndex        =   34
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txtInv3 
         DataField       =   "Inv3"
         DataSource      =   "mbr"
         Height          =   285
         Left            =   120
         TabIndex        =   33
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtInv3q 
         DataField       =   "Inv3q"
         DataSource      =   "mbr"
         Height          =   285
         Left            =   720
         TabIndex        =   32
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtInv4 
         DataField       =   "Inv4"
         DataSource      =   "mbr"
         Height          =   285
         Left            =   1320
         TabIndex        =   31
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox txtInv4q 
         DataField       =   "Inv4q"
         DataSource      =   "mbr"
         Height          =   285
         Left            =   1920
         TabIndex        =   30
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox txtInv5 
         DataField       =   "Inv5"
         DataSource      =   "mbr"
         Height          =   285
         Left            =   1320
         TabIndex        =   29
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txtInv5q 
         DataField       =   "Inv5q"
         DataSource      =   "mbr"
         Height          =   285
         Left            =   1920
         TabIndex        =   28
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txtInv6 
         DataField       =   "Inv6"
         DataSource      =   "mbr"
         Height          =   285
         Left            =   1320
         TabIndex        =   27
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtInv6q 
         DataField       =   "Inv6q"
         DataSource      =   "mbr"
         Height          =   285
         Left            =   1920
         TabIndex        =   26
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtGold 
         DataField       =   "Gold"
         DataSource      =   "mbr"
         Height          =   285
         Left            =   1320
         TabIndex        =   25
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtHp 
         DataField       =   "TotalHP"
         DataSource      =   "mbr"
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox txtMana 
         DataField       =   "TotalMana"
         DataSource      =   "mbr"
         Height          =   285
         Left            =   720
         TabIndex        =   23
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox txtHPLeft 
         DataField       =   "HPLeft"
         DataSource      =   "mbr"
         Height          =   285
         Left            =   1320
         TabIndex        =   22
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox txtManaLeft 
         DataField       =   "ManaLeft"
         DataSource      =   "mbr"
         Height          =   285
         Left            =   1920
         TabIndex        =   21
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox TxtExp 
         DataField       =   "Experence"
         DataSource      =   "mbr"
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Top             =   2880
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdItems 
      Caption         =   "Items"
      Height          =   255
      Left            =   2760
      TabIndex        =   18
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdturnl 
      Caption         =   "Turn >"
      Height          =   495
      Left            =   2040
      TabIndex        =   17
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton cmdturnr 
      Caption         =   "< Turn"
      Height          =   495
      Left            =   1440
      TabIndex        =   16
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton cmdGoAlleg 
      Caption         =   "&Allegria"
      Height          =   255
      Left            =   2760
      TabIndex        =   13
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "System"
      Height          =   1335
      Left            =   4080
      TabIndex        =   10
      Top             =   3240
      Width           =   1215
      Begin VB.CheckBox chkServtxt 
         Caption         =   "SText"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkWhisp 
         Caption         =   "Whispers"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   975
      End
      Begin VB.CheckBox chkFollow 
         Caption         =   "Follow"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   975
      End
      Begin VB.CheckBox chkServCode 
         Caption         =   "SCode"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.CommandButton cmduse 
      Caption         =   "&Use"
      Height          =   495
      Left            =   2040
      TabIndex        =   9
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton cmdWho 
      Caption         =   "&Who"
      Height          =   495
      Left            =   1440
      TabIndex        =   8
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton cmdSE 
      Caption         =   "SE"
      Height          =   495
      Left            =   720
      TabIndex        =   7
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton cmdSW 
      Caption         =   "SW"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton cmdNE 
      Caption         =   "NE"
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton cmdNW 
      Caption         =   "NW"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "&Disconnect"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Timer StayOnline 
      Interval        =   60000
      Left            =   120
      Top             =   600
   End
   Begin MSWinsockLib.Winsock sckFurc 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   5175
   End
   Begin VB.TextBox txtFromFurc 
      Height          =   2655
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frmBot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lastwalk As String
Dim whatwalk As String
Public Minute As Integer
Public Desc As String
Public Connected As Boolean
Public Doing As String
Public itemnum As Integer
Public Place As Integer

'Bot Settings
Const BotName = "Mesto"
Const BotPass = "0519aa"
Const descrip = "Figmit"
Const ColorCode = "! G2+88888!#!!#!"

Private Sub chkServCode_Click()
If chkServCode = 1 Then
chkServtxt = 2
chkServtxt.Enabled = False
End If
If chkServCode = 0 Then
chkServtxt = 1
chkServtxt.Enabled = True
End If
End Sub

Private Sub cmdGoAlleg_Click()
chkWhisp.Value = 0
chkFollow.Value = 0
sckFurc.SendData "goalleg" & vbLf
End Sub

Private Sub cmdItems_Click()
frmItems.Show
End Sub

Private Sub cmdNE_Click()
sckFurc.SendData "m 9" & vbLf
End Sub
Private Sub cmdNW_Click()
sckFurc.SendData "m 7" & vbLf
End Sub
Private Sub cmdSE_Click()
sckFurc.SendData "m 3" & vbLf
End Sub
Private Sub cmdSW_Click()
sckFurc.SendData "m 1" & vbLf
End Sub
Private Sub cmdturnl_Click()
sckFurc.SendData ">" & vbLf
End Sub
Private Sub cmdturnr_Click()
sckFurc.SendData "<" & vbLf
End Sub
Private Sub cmduse_Click()
sckFurc.SendData "use" & vbLf
End Sub
Private Sub cmdWho_Click()
sckFurc.SendData "who" & vbLf
End Sub

Sub Form_Load()
Minute = 0
Desc = descrip & " [Uptime: 0 Minute(s)]"
End Sub
Private Sub cmdConnect_Click()
If Connected = False Then
sckFurc.RemoteHost = "64.191.51.88"
sckFurc.RemotePort = "6000"
sckFurc.Connect
Connected = True
lastwalk = "none"
End If
End Sub
Private Sub cmdDisconnect_Click()
If Connected = True Then
sckFurc.Close
Connected = False
End If
End Sub

Private Sub sckFurc_DataArrival(ByVal bytesTotal As Long)
Dim s As String
sckFurc.GetData s
X = Split(s, vbLf)
For r = 0 To UBound(X) - 1
RealText X(r)
Next
End Sub
Sub RealText(Txt)
On Error Resume Next
If chkServtxt.Value = Checked Or chkServtxt.Enabled = False Then
If chkServCode.Value = Checked Then txtFromFurc = txtFromFurc & Txt & vbCrLf
If chkServCode.Value = Unchecked Then
If Left(Txt, 1) = "(" Then txtFromFurc = txtFromFurc & Right(Txt, Len(Txt) - 1) & vbCrLf
End If
End If
If Txt = "END" Then sckFurc.SendData "connect " & BotName & " " & BotPass & vbLf & "color " & ColorCode & vbLf & "desc " & Desc & vbLf
If Txt = "]ccmarbled.pcx" Then
sckFurc.SendData "vascodagama" & vbLf
End If

If Txt = "(Control of this dream is now being shared with you." Then
sckFurc.SendData "use" & vbLf & "m 7" & vbLf & "m 7" & vbLf & "m 7" & vbLf & ">" & vbLf & ">" & vbLf
chkWhisp.Value = 1
chkFollow.Value = 0
End If


If chkWhisp.Value = Checked Then
If Left(Txt, 3) = "([ " And Right(Txt, 10) = " to you. ]" Then
    tmsg = Split(Txt, " whispers, " & Chr(34), 2)
    furre = Right(tmsg(0), Len(tmsg(0)) - 3)
    NMsg = Left(tmsg(1), Len(tmsg(1)) - 11)
    Msg = LCase(NMsg)
    DoWhisper furre, Msg
End If
End If 'chkWhisp

'make the bot follow its owner
If frmBot.chkFollow.Value = Checked Then

If Left(Txt, 11) = "/!!8+6<<<<<" Then
        frl = Mid(Txt, 17, Len(Txt) - 0)
        whatwalk = Mid(frl, 1, Len(frl) - 4)
        'whatwalk = LCase(wwalk)
    dowalk whatwalk, lastwalk
End If
End If 'chkFollow

If Txt = "([#] signup" Then
sckFurc.SendData "l  ~ 0" & vbLf
Doing = "signup"
End If

If Left(Txt, 8) = "([#] do " Then
tmsg = Split(Txt, " # ", 7)

itemnum = tmsg(3)
Place = tmsg(2)

If Place = 1 And tmsg(1) = "buy" Then look = " r /"
If Place = 1 And tmsg(1) = "sell" Then look = " r /"
If Place = 1 And tmsg(1) = "fish" Then look = " % B"

sckFurc.SendData "l " & look & vbLf
Doing = tmsg(1)

End If

If Left(Txt, 10) = "((You see " Then
    furre = Mid(Txt, 11, Len(Txt) - 12)
    
    If Doing = "fish" Then
       dofish Place, furre
       Place = ""
       itemnum = ""
       
    ElseIf Doing = "signup" Then
       dosignup furre
       
    ElseIf Doing = "sell" Then
       dostore furre, Place, itemnum
       Place = ""
       itemnum = ""

    ElseIf Doing = "buy" Then
       dobuy furre, Place, itemnum
       Place = ""
       itemnum = ""

    End If
    
    Doing = ""
End If

End Sub

Sub dosignup(furre)


mbr.Recordset.MoveFirst
Do Until txtName.Text = furre Or mbr.Recordset.EOF
mbr.Recordset.MoveNext
Loop

If txtName.Text <> furre Then
mbr.Recordset.AddNew
txtName.Text = furre
txtDate.Text = Now
txtlvl.Text = 1
txtInv1q.Text = 0
txtInv2q.Text = 0
txtInv3q.Text = 0
txtInv4q.Text = 0
txtInv5q.Text = 0
txtInv6q.Text = 0
txtGold.Text = 0
txtHp = 25
txtMana = 15
txtHPLeft = 25
txtManaLeft = 15
TxtExp = 0
txtWeapon = 0
txtEquip1 = 0
txtEquip2 = 0
txtEquip3 = 0
txtEquip4 = 0

sckFurc.SendData Chr(34) & "signup done" & vbLf

Else

sckFurc.SendData Chr(34) & "signup member" & vbLf

End If

End Sub

Sub dofish(num, furre)

Dim catch
Randomize Timer
catch = Int((20 * Rnd) + 1)

If catch > 5 Then


mbr.Recordset.MoveFirst
Do Until txtName.Text = furre Or mbr.Recordset.EOF
mbr.Recordset.MoveNext
Loop

If txtName.Text <> furre Then

    sckFurc.SendData Chr(34) & "fish " & num & " yes nomember" & vbLf

Else

    If catch > 5 And catch < 11 Then
    additem 1, furre
        sckFurc.SendData Chr(34) & "fish " & num & " yes" & vbLf
    ElseIf catch > 10 And catch < 16 Then
    additem 2, furre
        sckFurc.SendData Chr(34) & "fish " & num & " something" & vbLf
    ElseIf catch > 15 And catch < 21 Then
    additem 3, furre
        sckFurc.SendData Chr(34) & "fish " & num & " something" & vbLf
    End If

End If


Else
sckFurc.SendData Chr(34) & "fish " & num & " no" & vbLf
End If

End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub StayOnline_Timer()
If Connected = True Then
Minute = Minute + 1
sckFurc.SendData "desc " & descrip & " [Uptime: " & Minute & " Minute(s)]" & vbLf
End If
End Sub


Private Sub txtFromFurc_Change()
txtFromFurc.SelStart = Len(txtFromFurc)
If Len(txtFromFurc) > 10000 Then txtFromFurc = Right(txtFromFurc, 9000)
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    sckFurc.SendData txtSend & vbLf
    txtSend = ""
    KeyAscii = 0
End If
End Sub
