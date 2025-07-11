Attribute VB_Name = "Else"
Option Explicit

Public Const MaxBot = 2

Dim Sign As String
Dim hit
Public Minute() As Integer
Public onet
Public twot

Public Type myBot
Name As String
Pass As String
Desc As String
Color As String
LastWalk As String
whatwalk As String
Minute As Integer
TriggerA As String
TriggerB As String
LookA As String
LookB As String
End Type

Public bot(0 To MaxBot) As myBot
