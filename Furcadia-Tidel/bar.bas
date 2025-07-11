Attribute VB_Name = "bar"
Sub doserve(Furre, Order, Index As Integer)

If Order = "beer" Then frmBoT.sckFurc(Index).SendData ":Pops the top on an ice cold Beer and sends it down the bar to " & Furre & "." & vbLf
If Order = "rootbeer" Then frmBoT.sckFurc(Index).SendData ":opens a bottle of A" & Chr(38) & "W RootBeer and hands it to " & Furre & "." & vbLf
If Order = "hamburger" Then frmBoT.sckFurc(Index).SendData ":Frys up a patty on the grill. Slaps lots of veggies and sauses on it to make it look biger. Puts it in a basket full of soggy frys and hands it to " & Furre & "." & vbLf
If Order = "hotdog" Then frmBoT.sckFurc(Index).SendData ":Pulls the oldest Hotdog off the turning hotdog cooker thing slips it into a dryed out bun and hands it to " & Furre & "." & vbLf

End Sub
