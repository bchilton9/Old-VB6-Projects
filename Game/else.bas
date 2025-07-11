Attribute VB_Name = "else"
Option Explicit

Public Declare Function TextOut& Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long)
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function BitBlt& Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWdith As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long)

Public Const srcCopy = &HCC0020
Public Const srcAnd = &H8800C6
Public Const srcPaint = &HEE0086
Public Const srcInvert = &H660046
Public Const srcErase = &H440328

'Client
Public InGame As Boolean
Public Speed As Byte
Public pIndex As Byte
Public CanWalk As Boolean
Public pChar As String * 1
Public pEnd As String * 1

Public Const Dir_Up = 0
Public Const Dir_Down = 2
Public Const Dir_Left = 3
Public Const Dir_Right = 1

Public Const serv = "24.8.47.166"
Public Const porta = 7171
Public Const portb = 7172

'Editor
Public MapNum As Long
Public LocalMap As MapRec

'Server/Client/Editor
Public Const MaxVar = 150
Public Const MaxClasses = 34
Public Const MaxPlayers = 50
Public Const MaxMaps = 30
Public Const MapX = 26
Public Const MapY = 13

Public Type ClassRec
  Name As String
  HP As Long
  MP As Long
  Str As Long
  Def As Long
End Type

Public Type playerinvitem
invitem(1 To 44) As Integer
End Type

Public Type Character
  Num As Long
  Name As String * 15
  Password As String * 15
  Access As Byte
  Time As Long
  HP As Integer
  MP As Integer
  Exp As Long
  Level As Byte
  Class As Byte
  Weapon As Byte
  Armor As Byte
  Shield As Byte
  Helmet As Byte
  map As Integer
  X As Integer
  Y As Integer
  xo As Integer
  yo As Integer
  d As Byte
  Walking As Boolean
  location As Integer
  inv As playerinvitem
  mask As Integer
  maskstep As Integer
  Container As Integer
  Height As Integer
  offx As Integer
  offy As Integer
  Set As Integer
End Type


Public Type MapTileRec
  TileX As Byte
  TileY As Byte
  Fringe As Byte
  Attrib As Byte
  Data1 As Byte
  Data2 As Byte
  Data3 As Byte
  objectoverX As Integer
  objectoverY As Integer
  objectX As Integer
  objectY As Integer
  npcX As Integer
  npcY As Integer
  Walk As Integer
End Type

Public Type MapNpcRec
  Npc As Long
  HP As Integer
  MP As Integer
  X As Byte
  Y As Byte
  xo As Integer
  yo As Integer
  d As Byte
  Walk As Integer
End Type

Public Type MapRec
  Name As String * 15
  Music As Byte
  Up As Integer
  Down As Integer
  Left As Integer
  Right As Integer
  Tile(0 To MapX, 0 To MapY) As MapTileRec
  Npc(0 To 4) As MapNpcRec
End Type

Public Type ItemRec
  Name As String * 15
  Cost As Long
  type As Byte
  Data1 As Byte
  Data2 As Byte
  Data3 As Byte
End Type

Public Type NpcRec
  Name As String * 15
  Move As Boolean
  pic As Byte
  HP As Integer
  MP As Integer
  Str As Byte
  Def As Byte
  Target As Byte
  type As Byte
  Data1 As Byte
  Data2 As Byte
  Data3 As Byte
  Data4 As Byte
  Data5 As Byte
End Type
  

Public Class(0 To MaxClasses) As ClassRec
Public Player(0 To MaxPlayers) As Character
Public map(0 To MaxMaps) As MapRec
Public Item(0 To MaxVar) As ItemRec
Public Npc(0 To MaxVar) As NpcRec

Public Sub SetParce()
  pChar = "#"
  pEnd = Chr(237)
End Sub
