<%
'-------------------------------------------------------------------------------
' Function Modifications - (MOST RECENT FIRST - ONLY FUNCTIONS FOR DATA DISPLAY TABLE)
'-------------------------------------------------------------------------------
' 02/24/05 - MDK - Created Initial File
'
'-------------------------------------------------------------------------------
' PURPOSE
'-------------------------------------------------------------------------------
' A function to draw a data-display table like data under an input box, for data
'  input.
'
'-------------------------------------------------------------------------------
' Instructions:
'-------------------------------------------------------------------------------
' Data is seperated from content for the purpose of using the same data set for
'  multiple input boxes.  DrawAutoComplete() data must be called first, then
'  DrawAutoCompleteBox() may be called as many times as needed on that data set.
'
' Note that ordering the recordset is irrelevant as the javascript auto-orders
'  all data
'
'-------------------------------------------------------------------------------
' SAMPLE EXECUTION:
'-------------------------------------------------------------------------------
' Data Set:
'
'   DataSetNm    = "ACP1"
'   ACPstrSQL    = "SELECT field1 [product name], field2 [sku] FROM ..."
'   fieldWidths  = array(65, 20, 15, 0)
'   textField    = "product name"
'   valField     = "p_key"
'   AutoComplete = "product name;sku"
'
' Optional Vars
'   width        = 300 ' if width is not set, will match input box's width.
'   maxHeight    = 150 ' default = 150
'   columnStyles = array(_
'                  "background-color:#ffffff;font-family:Tahoma;font-size:10px;overflow:hidden;padding:2px;", _
'                  "background-color:#ffffff;font-family:Tahoma;font-size:10px;overflow:hidden;padding:2px;", _
'                  "background-color:#ffffff;font-family:Tahoma;font-size:10px;overflow:hidden;padding:2px;", _
'                  ""                                                                                         _
'                  )
'   hoverStyle   = "background-color:#3366cc;color:#FFFFFF;"
'
'   DrawAutoCompleteData()
'
'
'
' Input Box:
'
'   ' titleStyle & contentStyle are only used on the first input box call.
'   titleStyle   = "background-color:#EECC99;border:1px solid #000000;font-family:Tahoma;font-size:10px;font-weight:bolx;padding:2px;"
'   contentStyle = "background-color:#FFFFFF;border:1px solid #000000;border-top:0px;"
'
'   inputID      = "complete"
'   inputNm      = "complete"
'   inputStyle   = "font-family:Tahoma;font-size:10px;width:200px;"
'   DataSet      = "ACP1"
'
'   DrawAutoCompleteBox()

DrawnACPInclude = false
nPrep           = ""
function DrawAutoCompleteData(byref oRs)
	' Begin Set Defaults
	if maxHeight    = "" then maxHeight    = 150
	if hoverStyle   = "" then hoverStyle   = "background-color:#3366cc;color:#FFFFFF;"
	if columnStyle  = "" then
		columnStyle  = array()
		redim columnStyle(ubound(fieldWidths)+1)
		for i = 0 to ubound(columnStyle)
			columnStyle(i) = "background-color:#ffffff;font-family:Tahoma;font-size:10px;overflow:hidden;padding:2px;"
		next
	end if
	' End Set Defaults

	nPrep    = nPrep & DataSetNm & ";"
	set nOut = new FastString

	nOut() = "<script>var " & DataSetNm & "={Nm:" & JSQ(DataSetNm)
	if width <> "" then nOut() = ",Width:" & clng(width)
	nOut() = ",maxHeight:" & clng(maxHeight)
	nOut() = ",titleStyle:" & JSQ(titleStyle)
	nOut() = ",contentStyle:" & JSQ(contentStyle)
	nOut() = ",Styles:["
	for i = 0 to ubound(columnStyle)
		nOut() = JSQ(columnStyle(i)) & cIf(i<ubound(columnStyle), ",", "")
	next
	nOut() = "]"
	nOut() = ",HLStyle:" & JSQ(HoverStyle)
	nOut() = ",fieldW:["
	for i = 0 to ubound(fieldWidths)
		nOut() = clng(fieldWidths(i)) & cIf(i<ubound(fieldWidths), ",", "")
	next
	nOut() = "]"

	tVal = 0
	vOf  = 0
	acp  = ""
	flds = ""
	ord  = 0

	i      = 0
	for each field in oRs.fields
		if lcase(field.name) = lcase(textField) then tVal = i
		if lcase(field.name) = lcase(valField)  then vOf  = i
		if lcase(field.name) = lcase(OrderBy)   then ord  = i
		if instr(1, ";" & AutoComplete & ";", field.name, 1)>0 then acp = acp & i & ","
		flds = flds & JSQ(field.name) & ","
		i      = i + 1
	next
	nOut() = ",Nms:[" & left(flds, len(flds)-1) & "]"
	nOut() = ",AutoCompleteOn:[" & left(acp, len(acp)-1) & "]"
	nOut() = ",TextVal:" & tVal
	nOut() = ",ValueOf:" & vOf
	if OrderBy <> "" then nOut() = ",OrderBy:" & ord

	j = 0
	nOut() = ",fields:["
	for each field in split(flds, ",")
		j = j + 1
		if field <> ""  then
			oRs.moveFirst()
			nOut() = "["
			do while not oRs.eof
				nOut() = JSQ(oRs(replace(Field, "'", "")))
				oRs.movenext()
				if not oRs.eof then nOut() = ","
			loop
			nOut() = "]" & cIf(j<i, ",", "")
		end if
	next
	nOut() = "]"
	nOut() = "}</script>"

	response.write nOut()
end function

function DrawAutoCompleteBox()
	if not DrawnACPInclude then
		if titleStyle   = "" then titleStyle   = "background-color:#EECC99;border:1px solid #000000;font-family:Tahoma;font-size:10px;font-weight:bolx;padding:2px;"
		if contentStyle = "" then contentStyle = "background-color:#FFFFFF;border:1px solid #000000;border-top:0px;"

		response.write "<div id=""titleDiv"" style=""" & TitleStyle & ";position:absolute;top:0px;left:0px;visibility:hidden;overflow:auto;z-Index:5;"" onmouseout=""offComplete();"" onmousemove=""onComplete();""></div><div id=""completeDiv"" hC=""0"" style=""" & ContentStyle & "position:absolute;top:0px;left:0px;visibility:hidden;overflow:auto;"" onmouseout=""offComplete();"" onmousedown=""setTimeout('cS=true', 10);"" onmousemove=""onComplete();""></div><div id=""holderDiv""></div>"
		response.write "<script language=""Javascript"" src=""common_images/scripts/ACPDropDown.js""></script>"
		DrawnACPInclude = true
	end if

	response.write "<input id=""" & inputID & """ name=""" & inputNm & """ style=""" & inputStyle & ";position:relative;"" onkeydown=""keyComplete(event, this, " & DataSet & ")"" onkeyup=""lV(this, " & DataSet & ")"" onmouseover=""this.onkeyup();"" onmouseout=""offComplete();"" onmousemove=""onComplete();"" onkeyup=""keyComplete()"" onfocus=""this.onkeyup();"" onblur=""cS=false;offComplete(50);""><input id=""" & inputID & "Val"" name=""" & inputNm & "Val"" type=""hidden"">"

	for each dset in split(nPrep, ";")
		response.write "<script>prepACP(" & dset & ");</script>"
	next
end function

function cIf(a, b, c)
	if cbool(a) then
		if IsObject(b) then set cIf = b else cIf = b
	else
		if IsObject(c) then set cIf = c else cIf = c
	end if
end function

function JSQ(text)
	JSQ = "'" & replace(text & "", "'", "\'") & "'"
end function
%>