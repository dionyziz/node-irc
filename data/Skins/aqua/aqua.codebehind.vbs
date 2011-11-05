' Aqua Skin
' CodeBehind
'
Sub Skin_LoadingCompleted()
	Dim webdocBack
	Dim htmlDiv
	Dim htmlImg
	Dim htmlAcronym
	Dim htmlLink
	Dim htmlClose
	Dim htmlSendText
	Dim htmlTextOptions
	Dim i
	
	htmlDiv = "<div id='xnode_skin_aqua_pretype' style='position:absolute; left: 2px; bottom: 65px;'>"
	htmlAcronym = "<acronym title='Send Text'>"
	htmlLink = "<a href=""NodeScript://Call Node.txtSend_KeyDown(13, 0):Node.txtSend.SetFocus '"">"
	htmlImg = "<img src='" & App.Path & "/data/skins/aqua/pretype.jpg' border='0' />"
	htmlClose = "</a></acronym></div>"

	htmlSendText = htmlDiv & htmlAcronym & htmlLink & htmlImg & htmlClose
	
	htmlDiv = "<div id='xnode_skin_aqua_postoptions' style='position:absolute; right: 15px; bottom: 65px;'>"
	htmlAcronym = "<acronym title='Text Format'>"
	htmlLink = "<a href=""NodeScript://Call Node.imgMore_Click() '"">"
	htmlImg = "<img src='" & App.Path & "/data/skins/aqua/up.jpg' border='0' />"
	htmlClose = "</a></acronym></div>"

	htmlTextOptions = htmlDiv & htmlAcronym & htmlLink & htmlImg & htmlClose
	
	With Node
		Set webdocBack = .wbBack.Document
		webdocBack.body.insertAdjacentHTML "BeforeEnd", htmlSendText
		webdocBack.body.insertAdjacentHTML "BeforeEnd", htmlTextOptions
		.sBar.Font.Name = "Verdana"
		.fraMore.Width = 0
		.fraMore.Height = 0
		Set .tbText.ImageList = Nothing
		.ilTbText.ListImages.Remove 8
		.ilTbText.ListImages.Remove 7
		.ilTbText.ListImages.Remove 6
		.ilTbText.ListImages.Remove 5
		.ilTbText.ListImages.Remove 4
		.ilTbText.ListImages.Remove 3
		.ilTbText.ListImages.Remove 2
		.ilTbText.ListImages.Remove 1
		.ilTbText.ListImages.Add , "Aqua_Smiley", .Scripting_LoadPicture( App.Path & "/data/skins/aqua/smile.jpg" )
		.ilTbText.ListImages.Add , "Aqua_Bold", .Scripting_LoadPicture( App.Path & "/data/skins/aqua/bold.jpg" )
		.ilTbText.ListImages.Add , "Aqua_Italic", .Scripting_LoadPicture( App.Path & "/data/skins/aqua/italic.jpg" )
		.ilTbText.ListImages.Add , "Aqua_Underline", .Scripting_LoadPicture( App.Path & "/data/skins/aqua/underline.jpg" )
		.ilTbText.ListImages.Add , "Aqua_Color", .Scripting_LoadPicture( App.Path & "/data/skins/aqua/color.jpg" )
		.ilTbText.ListImages.Add , "Aqua_Picture", .Scripting_LoadPicture( App.Path & "/data/skins/aqua/picture.jpg" )
		.ilTbText.ListImages.Add , "Aqua_Hyperlink", .Scripting_LoadPicture( App.Path & "/data/skins/aqua/hyperlink.jpg" )
		.ilTbText.ListImages.Add , "Aqua_Multiline", .Scripting_LoadPicture( App.Path & "/data/skins/aqua/multilines.jpg" )
		Set .tbText.ImageList = .ilTbText
		.tbText.Buttons(1).Image = 1
		.tbText.Buttons(3).Image = 2
		.tbText.Buttons(4).Image = 3
		.tbText.Buttons(5).Image = 4
		.tbText.Buttons(7).Image = 5
		.tbText.Buttons(9).Image = 6
		.tbText.Buttons(10).Image = 7
		.tbText.Buttons(11).Image = 8
		
		.tbText.Width = 450
		.tbText.Height = 1000
		
		.Caption = "Node :: Aqua"
	End With
End Sub
Sub Skin_Resize()
	Dim MyWebDoc

	With Node
		If .ntsPanel.Left <> 50 Then
			.ntsPanel.Left = 50
		End If
		.fraPanel.Height = .fraPanel.Height - 20
		'If Not .webdocChanMain Is Nothing Then
		'	Set MyWebDoc = .webdocChanMain
		'	MyWebDoc.parentWindow.frameBorder = "no";
		'End If
	End With
End Sub