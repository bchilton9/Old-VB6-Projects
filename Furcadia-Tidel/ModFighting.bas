Attribute VB_Name = "ModFighting"
Sub dofdream(turn, Index As Integer)
On Error Resume Next
anum = Int((5 * Rnd) + 1)
If anum = 1 Then atk = "Punched"
If anum = 2 Then atk = "Kicked"
If anum = 3 Then atk = "Bashed"
If anum = 4 Then atk = "Stabed"
If anum = 5 Then atk = "Slashed"

If turn(Index).Text = 1 Then

    rnum = frmBoT.f1lvl(Index).Text * 2
    num = Int((rnum * Rnd) + 1)
    hit = num + frmBoT.f1weapon(Index).Text - frmBoT.f2armor(Index).Text
    
    If hit < 1 Then hit = 1
    miss = Int((rnum * Rnd) + 1)
    
    If miss = hit Then
        frmBoT.sckFurc(Index).SendData Chr(34) & "emit " & frmBoT.f1name(Index).Text & "'s Attack Missed." & vbLf
turn(Index).Text = "2"
    Else
       frmBoT.f2hp(Index).Text = frmBoT.f2hp(Index).Text - hit
       frmBoT.sckFurc(Index).SendData Chr(34) & "emit " & frmBoT.f1name(Index).Text & " " & atk & " " & frmBoT.f2name(Index).Text & " For " & hit & " Points of damage, leaveing them with " & frmBoT.f2hp(Index).Text & " hitpoint's remaining." & vbLf
       
       
       If frmBoT.f2hp(Index).Text < 1 Then
       dowin frmBoT.f1name(Index).Text, frmBoT.f2name(Index).Text, Index
       Else
       turn(Index).Text = "2"
       End If
    End If

ElseIf turn(Index).Text = 2 Then
    
    rnum = frmBoT.f2lvl(Index).Text * 2
    num = Int((rnum * Rnd) + 1)
    hit = num + frmBoT.f1weapon(Index).Text - frmBoT.f1armor(Index).Text
    
    If hit < 1 Then hit = 1
    miss = Int((rnum * Rnd) + 1)
    
    If miss = hit Then
        frmBoT.sckFurc(Index).SendData Chr(34) & "emit " & frmBoT.f2name(Index).Text & "'s Attack Missed." & vbLf
        turn(Index).Text = "1"
    Else
       frmBoT.f1hp(Index).Text = frmBoT.f1hp(Index).Text - hit
       frmBoT.sckFurc(Index).SendData Chr(34) & "emit " & frmBoT.f2name(Index).Text & " " & atk & " " & frmBoT.f1name(Index).Text & " For " & hit & " Points of damage, leaveing them with " & frmBoT.f1hp(Index).Text & " hitpoint's remaining." & vbLf

       If frmBoT.f1hp(Index).Text < 1 Then
       dowin frmBoT.f2name(Index).Text, frmBoT.f1name(Index).Text, Index
       Else
       turn(Index).Text = "1"
       End If
    End If
End If

End Sub

Sub dowin(win, lose, Index As Integer)
On Error Resume Next
Open "members.txt" For Input As #1
Input #1, fName, mnum
Do Until (fName = win) Or (EOF(1))
Input #1, fName, mnum
Loop
Close #1
If fName = win Then
Open "memfiles\" & mnum & ".txt" For Input As #1
Input #1, fName, mnum, lvl, Class, gold, xp, weap, armo, sone, stwo, sthe, sfor, sfiv
Close #1
If lvl < 10 Then
nxp = xp + 10
End If
If lvl > 10 Then
nxp = xp + 5
End If
ngold = gold + 5

If nxp >= 100 Then
    nxp = 0
    lvl = lvl + 1
    frmBoT.sckFurc(Index).SendData Chr(34) & "emitloud " & win & " has ganed a lvl." & vbLf
End If
Open "memfiles\" & mnum & ".txt" For Output As #1
Write #1, fName, mnum, lvl, Class, gold, nxp, weap, armo, sone, stwo, sthe, sfor, sfiv
Close #1
End If

Open "members.txt" For Input As #1
Input #1, fName, mnum
Do Until (fName = lose) Or (EOF(1))
Input #1, fName, mnum
Loop
Close #1
If fName = lose Then
Open "memfiles\" & mnum & ".txt" For Input As #1
Input #1, fName, mnum, lvl, Class, gold, xp, weap, armo, sone, stwo, sthe, sfor, sfiv
Close #1

If lvl < 10 Then
nxp = xp + 7
End If
If lvl > 10 Then
nxp = xp + 3
End If
If nxp >= 100 Then
    nxp = 0
    lvl = lvl + 1
    frmBoT.sckFurc(Index).SendData Chr(34) & "emitloud " & lose & " has ganed a lvl." & vbLf
End If
Open "memfiles\" & mnum & ".txt" For Output As #1
Write #1, fName, mnum, lvl, Class, gold, nxp, weap, armo, sone, stwo, sthe, sfor, sfiv
Close #1
End If

frmBoT.sckFurc(Index).SendData Chr(34) & "emitloud " & win & " has defeted " & lose & vbLf
frmBoT.reset Index
End Sub
