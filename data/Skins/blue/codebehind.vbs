'CodeBehind for Blue Skin
Sub Skin_Init()
	Node.wbStatus(0).Left = -100
	Node.wbStatus(0).Top = nmnuMain.Height + tsTabs.Height + 100
End Sub
Sub Skin_LoadingCompleted()
	With frmOptions
		.icLanguage.Font.Size = 10
		.fraLine.Left = 2500
		.fraLine.Appearance = 0
		.fraLine.BackColor = RGB(52, 74, 133)
		'Load .picCustom(1)
		'.picCustom(0).Visible = True
		'.picCustom(0).Appearance = 0
		'.picCustom(0).Picture = Scripting_LoadPicture(App.Path & "\data\skins\blue\options_back.jpg")
		'.picCustom(0).Left = .fraLine.Left + 30
		'.picCustom(0).Top = .fraLine.Top - 30
		'.picCustom(0).BorderStyle = 0
		'.picCustom(0).AutoSize = True
		'.picCustom(0).ZOrder 0
	End With
End Sub