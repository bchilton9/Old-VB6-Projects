Attribute VB_Name = "skinner"
Sub loadskin(mySkin As String)

'On Error GoTo noload

Open App.Path & "\Skins\" & mySkin & "\" & mySkin & ".txt" For Input As #1
Input #1, junk, Main_Back_Image
Input #1, junk, Send_Button_Image
Input #1, junk, Send_Button_Down_Image
Input #1, junk, Send_Button_X
Input #1, junk, Send_Button_Y
Input #1, junk, Chat_X
Input #1, junk, Chat_Y
Input #1, junk, Send_Chat_X
Input #1, junk, Send_Chat_Y
Input #1, junk, Buffer_X
Input #1, junk, Buffer_Y
Input #1, junk, Who_Button_Image
Input #1, junk, Who_Button_Down_Image
Input #1, junk, Who_Button_X
Input #1, junk, Who_Button_Y
Input #1, junk, Inventory_Button_Image
Input #1, junk, Inventory_Button_Down_Image
Input #1, junk, Inventory_Button_X
Input #1, junk, Inventory_Button_Y
Input #1, junk, Inventory_Background
Input #1, junk, Inventory_X
Input #1, junk, Inventory_Y
Input #1, junk, Inventory_X_Head
Input #1, junk, Inventory_Y_Head
Input #1, junk, Inventory_X_Neck
Input #1, junk, Inventory_Y_Neck
Input #1, junk, Inventory_X_Tunic
Input #1, junk, Inventory_Y_Tunic
Input #1, junk, Inventory_X_Belt
Input #1, junk, Inventory_Y_Belt
Input #1, junk, Inventory_X_Boots
Input #1, junk, Inventory_Y_Boots
Input #1, junk, Inventory_X_Earring
Input #1, junk, Inventory_Y_Earring
Input #1, junk, Inventory_X_Bracer
Input #1, junk, Inventory_Y_Bracer
Input #1, junk, Inventory_X_Sleaves
Input #1, junk, Inventory_Y_Sleaves
Input #1, junk, Inventory_X_Gloves
Input #1, junk, Inventory_Y_Gloves
Input #1, junk, Inventory_X_Legs
Input #1, junk, Inventory_Y_Legs
Input #1, junk, Inventory_X_Weapon
Input #1, junk, Inventory_Y_Weapon
Input #1, junk, Inventory_X_Shield
Input #1, junk, Inventory_Y_Shield
Input #1, junk, Inventory_Close_Image
Input #1, junk, Inventory_CloseOver_Image
Input #1, junk, Inventory_Close_X
Input #1, junk, Inventory_Close_Y
Input #1, junk, Inventory_Inspect_Image
Input #1, junk, Inventory_Inspectover_Image
Input #1, junk, Inventory_Inspect_X
Input #1, junk, Inventory_Inspect_Y
Input #1, junk, Inventory_Distory_Image
Input #1, junk, Inventory_Distoryover_Image
Input #1, junk, Inventory_Distory_X
Input #1, junk, Inventory_Distory_Y

Close #1

picPath = App.Path & "\Skins\" & mySkin & "\" & Main_Back_Image
frmMain.Picture = LoadPicture(picPath)

picPath = App.Path & "\Skins\" & mySkin & "\" & Send_Button_Image
frmMain.cmdSend.Picture = LoadPicture(picPath)
frmMain.cmdSend3.Picture = LoadPicture(picPath)
picPath = App.Path & "\Skins\" & mySkin & "\" & Send_Button_Down_Image
frmMain.cmdSend2.Picture = LoadPicture(picPath)

frmMain.cmdSend.Left = Send_Button_X
frmMain.cmdSend.Top = Send_Button_Y

frmMain.txtChat.Left = Chat_X
frmMain.txtChat.Top = Chat_Y

frmMain.txtSend.Left = Send_Chat_X
frmMain.txtSend.Top = Send_Chat_Y

frmMain.picBuffer.Left = Buffer_X
frmMain.picBuffer.Top = Buffer_Y

frmMain.frmSign.Left = (Buffer_X + 64)
frmMain.frmSign.Top = (Buffer_Y + 64)

frmMain.frmInventory.Left = Buffer_X
frmMain.frmInventory.Top = Buffer_Y

frmMain.frmBank.Left = (Buffer_X + 6855)
frmMain.frmBank.Top = Buffer_Y

frmMain.frmBlank.Left = (Buffer_X + 6855)
frmMain.frmBlank.Top = Buffer_Y

frmMain.frmStore.Left = (Buffer_X + 6855)
frmMain.frmStore.Top = Buffer_Y

frmMain.frmBattle.Left = Buffer_X
frmMain.frmBattle.Top = Buffer_Y

picPath = App.Path & "\Skins\" & mySkin & "\" & Who_Button_Image
frmMain.cmdWho.Picture = LoadPicture(picPath)
frmMain.cmdWho3.Picture = LoadPicture(picPath)
picPath = App.Path & "\Skins\" & mySkin & "\" & Who_Button_Down_Image
frmMain.cmdWho2.Picture = LoadPicture(picPath)

frmMain.cmdWho.Left = Who_Button_X
frmMain.cmdWho.Top = Who_Button_Y

picPath = App.Path & "\Skins\" & mySkin & "\" & Inventory_Button_Image
frmMain.cmgInv.Picture = LoadPicture(picPath)
frmMain.cmgInv3.Picture = LoadPicture(picPath)
picPath = App.Path & "\Skins\" & mySkin & "\" & Inventory_Button_Down_Image
frmMain.cmgInv2.Picture = LoadPicture(picPath)

frmMain.cmgInv.Left = Inventory_Button_X
frmMain.cmgInv.Top = Inventory_Button_Y

picPath = App.Path & "\Skins\" & mySkin & "\" & Inventory_Background
frmMain.picInvBack.Picture = LoadPicture(picPath)

frmMain.picEquip(13).Top = Inventory_X
frmMain.picEquip(13).Left = Inventory_Y
frmMain.picEquip(14).Top = Inventory_X
frmMain.picEquip(14).Left = Inventory_Y + 600
frmMain.picEquip(15).Top = Inventory_X
frmMain.picEquip(15).Left = Inventory_Y + 1200
frmMain.picEquip(16).Top = Inventory_X
frmMain.picEquip(16).Left = Inventory_Y + 1800
frmMain.picEquip(17).Top = Inventory_X + 600
frmMain.picEquip(17).Left = Inventory_Y
frmMain.picEquip(18).Top = Inventory_X + 600
frmMain.picEquip(18).Left = Inventory_Y + 600
frmMain.picEquip(19).Top = Inventory_X + 600
frmMain.picEquip(19).Left = Inventory_Y + 1200
frmMain.picEquip(20).Top = Inventory_X + 600
frmMain.picEquip(20).Left = Inventory_Y + 1800
frmMain.picEquip(21).Top = Inventory_X + 1200
frmMain.picEquip(21).Left = Inventory_Y
frmMain.picEquip(22).Top = Inventory_X + 1200
frmMain.picEquip(22).Left = Inventory_Y + 600
frmMain.picEquip(23).Top = Inventory_X + 1200
frmMain.picEquip(23).Left = Inventory_Y + 1200
frmMain.picEquip(24).Top = Inventory_X + 1200
frmMain.picEquip(24).Left = Inventory_Y + 1800
frmMain.picEquip(25).Top = Inventory_X + 1800
frmMain.picEquip(25).Left = Inventory_Y
frmMain.picEquip(26).Top = Inventory_X + 1800
frmMain.picEquip(26).Left = Inventory_Y + 600
frmMain.picEquip(27).Top = Inventory_X + 1800
frmMain.picEquip(27).Left = Inventory_Y + 1200
frmMain.picEquip(28).Top = Inventory_X + 1800
frmMain.picEquip(28).Left = Inventory_Y + 1800

frmMain.picEquip(6).Top = Inventory_X_Head
frmMain.picEquip(6).Left = Inventory_Y_Head
frmMain.picEquip(7).Top = Inventory_X_Neck
frmMain.picEquip(7).Left = Inventory_Y_Neck
frmMain.picEquip(5).Top = Inventory_X_Tunic
frmMain.picEquip(5).Left = Inventory_Y_Tunic
frmMain.picEquip(12).Top = Inventory_X_Belt
frmMain.picEquip(12).Left = Inventory_Y_Belt
frmMain.picEquip(9).Top = Inventory_X_Boots
frmMain.picEquip(9).Left = Inventory_Y_Boots
frmMain.picEquip(10).Top = Inventory_X_Earring
frmMain.picEquip(10).Left = Inventory_Y_Earring
frmMain.picEquip(8).Top = Inventory_X_Bracer
frmMain.picEquip(8).Left = Inventory_Y_Bracer
frmMain.picEquip(4).Top = Inventory_X_Sleaves
frmMain.picEquip(4).Left = Inventory_Y_Sleaves
frmMain.picEquip(11).Top = Inventory_X_Gloves
frmMain.picEquip(11).Left = Inventory_Y_Gloves
frmMain.picEquip(3).Top = Inventory_X_Legs
frmMain.picEquip(3).Left = Inventory_Y_Legs
frmMain.picEquip(1).Top = Inventory_X_Weapon
frmMain.picEquip(1).Left = Inventory_Y_Weapon
frmMain.picEquip(2).Top = Inventory_X_Shield
frmMain.picEquip(2).Left = Inventory_Y_Shield

picPath = App.Path & "\Skins\" & mySkin & "\" & Inventory_Close_Image
frmMain.cmdCloseInventory.Picture = LoadPicture(picPath)
picPath = App.Path & "\Skins\" & mySkin & "\" & Inventory_CloseOver_Image
frmMain.cmdCloseInventory2.Picture = LoadPicture(picPath)
picPath = App.Path & "\Skins\" & mySkin & "\" & Inventory_Close_Image
frmMain.cmdCloseInventory3.Picture = LoadPicture(picPath)
frmMain.cmdCloseInventory.Top = Inventory_Close_X
frmMain.cmdCloseInventory.Left = Inventory_Close_Y

picPath = App.Path & "\Skins\" & mySkin & "\" & Inventory_Inspect_Image
frmMain.cmdInspect.Picture = LoadPicture(picPath)
picPath = App.Path & "\Skins\" & mySkin & "\" & Inventory_Inspectover_Image
frmMain.cmdInspect2.Picture = LoadPicture(picPath)
picPath = App.Path & "\Skins\" & mySkin & "\" & Inventory_Inspect_Image
frmMain.cmdInspect3.Picture = LoadPicture(picPath)
frmMain.cmdInspect.Top = Inventory_Inspect_X
frmMain.cmdInspect.Left = Inventory_Inspect_Y

picPath = App.Path & "\Skins\" & mySkin & "\" & Inventory_Distory_Image
frmMain.cmdDistory.Picture = LoadPicture(picPath)
picPath = App.Path & "\Skins\" & mySkin & "\" & Inventory_Distoryover_Image
frmMain.cmdDistory2.Picture = LoadPicture(picPath)
picPath = App.Path & "\Skins\" & mySkin & "\" & Inventory_Distory_Image
frmMain.cmdDistory3.Picture = LoadPicture(picPath)
frmMain.cmdDistory.Top = Inventory_Distory_X
frmMain.cmdDistory.Left = Inventory_Distory_Y

Exit Sub

noload:
    Unload frmTCP
    Unload frmMain
msg = MsgBox("Unable to load skin file. Please check your skins and try agine.", vbCritical, "Load Skin Error")
End
End Sub
