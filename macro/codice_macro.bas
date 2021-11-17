
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


' Fix all errors

Public Sub FixAll()

	' Check if it is a text document
	oCurrentController = ThisComponent.getCurrentController()	
	if not(oCurrentController.supportsService("com.sun.star.text.TextDocumentView")) then
		msgbox("Macro works only in text documents.")
		exit sub
	end if
	
	' Operate
	nTable = FixTableFormat()
	nFormula = FixFormulaFormat()
	nIndex = UpdateIndexes()

	' Feedback
	print nTable & " tables, " & nFormula & " formulas, and " & nIndex & " indexes fixed"
	
End Sub


' Fix table format

Private Function FixTableFormat() as Integer
    oTexttables = thisComponent.Texttables
    for i = 0 to oTexttables.count - 1
        oTexttable = oTexttables(i)
        oTexttable.split = false
        oTexttable.keeptogether = true
        oTexttable.repeatheadline = false
        oTexttable.horiorient = 6
        oTexttable.topmargin = 0       
        oTexttable.bottommargin = 0
        ' Fix border width of all table cells
        v = oTexttable.TableBorder
        
		x = v.TopLine
		if x.OuterLineWidth <> 0 then
			x.OuterLineWidth = 8
			v.TopLine = x
		end if

		x = v.LeftLine
		if x.OuterLineWidth <> 0 then
			x.OuterLineWidth = 8
			v.LeftLine = x
		end if

		x = v.RightLine
		if x.OuterLineWidth <> 0 then
			x.OuterLineWidth = 8
			v.RightLine = x
		end if

		x = v.BottomLine
		if x.OuterLineWidth <> 0 then
			x.OuterLineWidth = 8
			v.BottomLine = x
		end if
		
		x = v.VerticalLine
		if x.OuterLineWidth <> 0 then
			x.OuterLineWidth = 8
			v.VerticalLine = x
		end if

		x = v.HorizontalLine
		if x.OuterLineWidth <> 0 then
			x.OuterLineWidth = 8
			v.HorizontalLine = x
		end if

		oTexttable.TableBorder = v
	next i
	FixTableFormat() = oTexttables.count
End Function


' Fix format of all formulas

Private Function FixFormulaFormat() as Integer
    oEmbeddedObjects = thisComponent.EmbeddedObjects
    nFormula = 0
    for i = 0 to oEmbeddedObjects.count - 1
    	oEmbeddedObject = oEmbeddedObjects.getbyIndex(i)
		oModel = oEmbeddedObject.Model  ' might be empty
		if isEmpty(oModel) then continue 
		' Is it a formula?
		if oModel.supportsService("com.sun.star.formula.FormulaProperties") then
			' Common settings
			oModel.LeftMargin = 0
			oModel.RightMargin = 0
			oModel.TopMargin = 0
			oModel.BottomMargin = 0
			oModel.FontTextIsBold = false
			oModel.FontVariablesIsBold = false
			oModel.FontFunctionsIsBold = false
			oModel.FontNumbersIsBold = false
			oModel.FontTextIsItalic = false
			oModel.FontVariablesIsItalic = false
			oModel.FontFunctionsIsItalic = false
			oModel.FontNumbersIsItalic = false
			if oModel.BaseFontHeight = 9 then  ' Probably in a table
				oModel.FontNameText = "Liberation Sans"
				oModel.FontNameFunctions = "Liberation Sans"
				oModel.FontNameVariables = "Liberation Sans"
				oModel.FontNameNumbers = "Liberation Sans"			
				oModel.CustomFontNameSerif= "Liberation Serif"
				oModel.CustomFontNameSans= "Liberation Sans"
				oModel.CustomFontNameFixed= "Liberation Mono"
			else  ' Probably out of a table
				oModel.BaseFontHeight = 11
				oModel.FontNameText = "Liberation Serif"
				oModel.FontNameFunctions = "Liberation Serif"
				oModel.FontNameVariables = "Liberation Serif"
				oModel.FontNameNumbers = "Liberation Serif"
				oModel.CustomFontNameSerif= "Liberation Serif"
				oModel.CustomFontNameSans= "Liberation Sans"
				oModel.CustomFontNameFixed= "Liberation Mono"
			end if
			' Update and count
			oXCOEO = oEmbeddedObject.ExtendedControlOverEmbeddedObject
			oXCOEO.update()
			nFormula = nFormula + 1
		end if
	next i
	FixFormulaFormat() = nFormula
end function


' Update indexes

Private Function UpdateIndexes() as Integer
	oIndexes = ThisComponent.getDocumentIndexes()
	for i = 0 to oIndexes.getCount() - 1
		oIndexes(i).update
	next i
	UpdateIndexes() = oIndexes.count
end function


' Find common typing errors

Public Sub FindTextError()

	' Check type of document
	oCurrentController = ThisComponent.getCurrentController()
	if not(oCurrentController.supportsService("com.sun.star.text.TextDocumentView")) then
		msgbox("Macro works only in text documents.")
		exit sub
	end if

	' Get current document and view cursor
	oDoc = ThisComponent
	oCursor= oCurrentController.getViewCursor
	
	' Define error regexp array
	' Array of error regex expressions:
	sAll      = "[[:alnum:]\-\.\,\:\°\(\[\{\“\€\±\)\]\}\”\%\!\?\;\+\=\/\·\×\÷\−\≤\≥\≠\≈\@\’\–\—\u003C\u003E]" ' Unicode: <, >
	sAllAlpha = "[[:alpha:]\-\.\,\:\°\(\[\{\“\€\±\)\]\}\”\%\!\?\;\+\=\/\·\×\÷\−\≤\≥\≠\≈\@\’\–\—\u003C\u003E]"
	sAllPunct = "[\.\,\:\”\!\?\;]"
	sBefore   = "[\(\[\{\“\€\±\§]"             ' Space only before
	sAfter    = "[\.\!\?\;]"                   ' Space only after
	sBoth     = "[\+\=\·\/\u003C\u003E\×\÷\−\≤\≥\≠\≈]" ' Space before and after, nb <, >
	sNone     = "[\°\@\’\–\—]"                 ' No space
	
	sEmptyLine = "[\n\r]{2,}"                 ' Empty line FIXME
	sSpace1  = "[:space:]{2,}" 		          ' No multiple spaces
	sSpace2  = "[:space:]+$"                  ' No spaces at the end of paragraph
	sUnused  = "[\|\\\*\#\_\'\‘\«\»\^\u0022]+"  ' Unused chars, char(34) = "


	sMinus1  = sAll & "\-[:space:]"           ' No: text-_
	sMinus2  = "[:space:]\-" & sAllAlpha      ' No: _-text
	sComma1  = "\," & sAllAlpha               ' No: ,text Allowed: ,123
	sComma2  = "[:space:]\,"                  ' No: _,
	sColon1  = "[:space:]\:"                  ' No: _:
	sColon2  = "\:" & sAllAlpha               ' No: :text Allowed: 9001:2012
	sBraketClosed1 = "[\)\]\}\”\%]" & "[[:alnum:]\-\°\(\[\{\“\€\±\%\+\=\/\·\×\÷\−\≤\≥\≠\≈\@\’\–\—\u003C\u003E]"
	sWords = "ovvero|superiore|inferiore|capitol|paragraf|FIXME|di prestazione|punt|cartellonistica|livello I"
	
	'Fix /: http://www.bsigroup.com/
	'Fix 30'
		
	oErrors = Array( _
		sEmptyLine, sSpace1, sSpace2, sUnused, _
		sAll & sBefore, sBefore & "[:space:]", _
		sAfter & sAll, "[:space:]" & sAfter, _
		sAll & sBoth & sAll, _
		"[:space:]" & sNone, sNone & "[:space:]", _
		sMinus1, sMinus2, sComma1, sComma2, sColon1, sColon2, sBraketClosed1, _
		sWords, _
	)
		
	' Get search descriptor
	oFind = oDoc.createSearchDescriptor()
	with oFind
		.SearchRegularExpression = true
		.SearchCaseSensitive = false
	end with

	sError = Join(oErrors, "|")
	allChecked = False
	Do
		' Search for errors and select the next found
    	oFind.SearchString = sError
		oFound = oDoc.FindNext(oCursor, oFind)
		' Is there an error?
		If not IsEmpty(oFound) and not IsNull(oFound) Then
			oDoc.CurrentController.select(oFound)
			exit sub
		End If
		' All checked? Otherwise restart
		If allChecked Then exit do
		allChecked = True
		oCursor.gotoStart(false)
	Loop
	msgbox("No errors found")
End Sub

