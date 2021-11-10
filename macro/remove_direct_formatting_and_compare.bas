
Public sub CompareCurrentDocWith(filepath as String)
	' Compare current document with the called one

	dim oDocFrame as object, dispatcher as object
	dim PropVal(0) as new com.sun.star.beans.PropertyValue
	dim args(0) as new com.sun.star.beans.PropertyValue

	oDocFrame = ThisComponent.CurrentController.Frame
	
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

	PropVal(0).Name = "URL"
	PropVal(0).Value = convertToUrl(filepath)	
	dispatcher.executeDispatch(oDocFrame, ".uno:CompareDocuments", "", 0, PropVal())
	
	args(0).Name = "ShowTrackedChanges"
	args(0).Value = true
	dispatcher.executeDispatch(oDocFrame, ".uno:ShowTrackedChanges", "", 0, args())

end sub


Public Sub RmDirectFormatting()
	' Translate most frequent direct formatting to styles
	' then removes all direct formatting from the document

	' Check correct type of document
	oCurrentController = ThisComponent.getCurrentController()
	if not(oCurrentController.supportsService("com.sun.star.text.TextDocumentView")) then
		msgbox("Only for Writer documents.")
		exit sub
	end if

	Dim SearchAttributes(0) As New com.sun.star.beans.PropertyValue
	Dim SearchAttributes1(1) As New com.sun.star.beans.PropertyValue
	Dim SearchAttributes2(2) As New com.sun.star.beans.PropertyValue
	
	' fix char styles
	
	SearchAttributes(0).Name = "CharWeight"
	SearchAttributes(0).Value = com.sun.star.awt.FontWeight.BOLD
	NewCharStyleName = "Strong Emphasis"
	_ReplaceCharDirectFormatting(SearchAttributes, NewCharStyleName)

	SearchAttributes(0).Name = "CharPosture"
	SearchAttributes(0).Value = com.sun.star.awt.FontSlant.ITALIC
	NewCharStyleName = "Emphasis"
	_ReplaceCharDirectFormatting(SearchAttributes, NewCharStyleName)

	SearchAttributes2(0).Name = "CharAutoEscapement"
	SearchAttributes2(0).Value = true
	SearchAttributes2(1).Name = "CharEscapementHeight"
	SearchAttributes2(1).Value = 58
	SearchAttributes2(2).Name = "CharEscapement"
	SearchAttributes2(2).Value = 101
	NewCharStyleName = "Superscript"
	_ReplaceCharDirectFormatting(SearchAttributes2, NewCharStyleName)

	SearchAttributes2(0).Name = "CharAutoEscapement"
	SearchAttributes2(0).Value = true
	SearchAttributes2(1).Name = "CharEscapementHeight"
	SearchAttributes2(1).Value = 58
	SearchAttributes2(2).Name = "CharEscapement"
	SearchAttributes2(2).Value = -101
	NewCharStyleName = "Subscript"
	_ReplaceCharDirectFormatting(SearchAttributes2, NewCharStyleName)

	' fix paragraph styles

	SearchAttributes1(0).Name = "ParaAdjust"
	SearchAttributes1(0).Value = 3 ' com.sun.star.awt.ParagraphAdjust.CENTER
	SearchAttributes1(1).Name = "ParaStyleName"
	SearchAttributes1(1).Value = "Table Contents"
	NewParaStyleName = "Table Contents Center"
	_ReplaceParaDirectFormatting(SearchAttributes1, NewParaStyleName)

	SearchAttributes1(0).Name = "ParaAdjust"
	SearchAttributes1(0).Value = 2 ' com.sun.star.awt.ParagraphAdjust.BLOCK
	SearchAttributes1(1).Name = "ParaStyleName"
	SearchAttributes1(1).Value = "Table Contents"
	NewParaStyleName = "Table Contents Justify"
	_ReplaceParaDirectFormatting(SearchAttributes1, NewParaStyleName)
	
	' other?
	'TextPortion.setPropertyToDefault("CharFontName")
	'TextPortion.setPropertyToDefault("CharColor")
	'TextPortion.setPropertyToDefault("CharHeight")
	'TextPortion.setPropertyToDefault("CharUnderline")
    'TextPortion.setPropertyToDefault("CharWeight")
    'TextPortion.setPropertyToDefault("CharPosture")
	'TextPortion.setPropertyToDefault("CharBackColor")
	'TextPortion.setPropertyToDefault("CharEscapement")
	'TextPortion.setPropertyToDefault("CharEscapementHeight")
	'TextPortion.setPropertyToDefault("CharStrikeout")
	'TextPortion.setPropertyToDefault("CharUnderline")

	' reset all character and paragraph attributes
	oFrame   = oCurrentController.Frame
	oDispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	oDispatcher.executeDispatch(oFrame, ".uno:SelectAll", "", 0, Array())
	oDispatcher.executeDispatch(oFrame, ".uno:ResetAttributes", "", 0, Array())

End Sub


sub _ReplaceCharDirectFormatting(SearchAttributes as com.sun.star.beans.PropertyValue, NewCharStyleName as string)

	oDoc = ThisComponent
	oCurrentController = oDoc.getCurrentController()

	' get new search descriptor and configure it
	oFind = oDoc.createSearchDescriptor()
	with oFind
	    .SetSearchAttributes(SearchAttributes())
		.searchAll = True
	end with

	' treat oFound
    oFound = oDoc.findFirst(oFind)
    Do While Not IsNull(oFound)
        if oFound.getPropertyState(SearchAttributes(0).Name) = com.sun.star.beans.PropertyState.DIRECT_VALUE Then
        	'oFound.setPropertyToDefault(SearchAttributes(0).Name)
			oFound.CharStyleName = NewCharStyleName
		end if
        oFound = oDoc.findNext(oFound.End, oFind)
    Loop
    
end sub

sub _ReplaceParaDirectFormatting(SearchAttributes as com.sun.star.beans.PropertyValue, NewParaStyleName as string)

	oDoc = ThisComponent
	oCurrentController = oDoc.getCurrentController()

	' get new search descriptor and configure it
	oFind = oDoc.createSearchDescriptor()
	with oFind
	    .SetSearchAttributes(SearchAttributes())
		.SearchAll = True
	end with

	' treat oFound
    oFound = oDoc.findFirst(oFind)
    Do While Not IsNull(oFound)
		oFound.ParaStyleName = NewParaStyleName
        oFound = oDoc.findNext(oFound.End, oFind)
    Loop
   
end sub

