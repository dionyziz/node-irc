'CodeBehind for Blue Skin
Sub Skin_Init()
	Node.wbStatus(0).Left = -100
End Sub
Sub Skin_LoadingCompleted()
	With frmOptions
		.icLanguage.Font.Size = 10
		.fraLine.Left = 2500
		.fraLine.Appearance = 0
		.fraLine.BackColor = RGB(63, 137, 230) 
		.fraLine.BorderStyle = 0
		.fraLine.Width = 6015
		.fraLine.Left = 2530
		
		.Scripting_LoadImage(1)
		.picCustom(1).Visible = True
		.picCustom(1).Appearance = 0
		.picCustom(1).Picture = Scripting_LoadPicture(App.Path & "\data\skins\nodexp\options_top.jpg")
		.picCustom(1).Left = .fraLine.Left + 30
		.picCustom(1).Top = 0
		.picCustom(1).BorderStyle = 0
		.picCustom(1).AutoSize = True
		.picCustom(1).ZOrder 0
		
		.Scripting_LoadImage(2)
		.picCustom(2).Visible = True
		.picCustom(2).Appearance = 0
		.picCustom(2).Picture = Scripting_LoadPicture(App.Path & "\data\skins\nodexp\options_right.jpg")
		.picCustom(2).Left = .Width - 300
		.picCustom(2).Top = .picCustom(1).Height + .picCustom(1).Top
		.picCustom(2).BorderStyle = 0
		.picCustom(2).AutoSize = True
		.picCustom(2).ZOrder 0
		
		.Scripting_LoadImage(3)
		.picCustom(3).Visible = True
		.picCustom(3).Appearance = 0
		.picCustom(3).Picture = Scripting_LoadPicture(App.Path & "\data\skins\nodexp\options_bottom.jpg")
		.picCustom(3).AutoSize = True
		.picCustom(3).Left = .fraLine.Left + 30
		.picCustom(3).Top = .fraLine.Top
		.picCustom(3).BorderStyle = 0

		.fraLine.ZOrder 0
		.cmdOK.ZOrder 0
		.cmdCancel.ZOrder 0
		.cmdApply.ZOrder 0
		.picCustom(3).ZOrder 0
		
		.Scripting_LoadImage(4)
		.picCustom(4).Visible = True
		.picCustom(4).Appearance = 0
		.picCustom(4).Picture = Scripting_LoadPicture(App.Path & "\data\skins\nodexp\options_bottom_left.jpg")
		.picCustom(4).AutoSize = True
		.picCustom(4).Left = .fraLine.Left
		.picCustom(4).Top = .fraLine.Top
		.picCustom(4).BorderStyle = 0
		.picCustom(4).ZOrder 0		
	End With
End Sub