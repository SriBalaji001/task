<!-- #include file = "LIB_AutoCompleteBox_022405.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>

<BODY bgcolor="#E6E6E6">
<%
   DataSetNm    = "SC"
   ACPstrSQL    = "SELECT city as [City], zipcode as [Zip Code] FROM zip_code_map WHERE statecode='SC' ORDER BY City"
   fieldWidths  = array(75, 25)
'	Width        = 500	' uncomment to set width
	OrderBy      = "City"
   textField    = "City"
   valField     = "Zip Code"
   AutoComplete = "City;Zip Code"

	' Open Recordset Here!
	' set oRs = ....
   DrawAutoCompleteData(oRs)


   DataSetNm    = "NC"
   ACPstrSQL    = "SELECT city as [City], zipcode as [Zip Code] FROM zip_code_map WHERE statecode='NC' ORDER BY City"

	' Open Recordset Here!
	' set oRs = ....
   DrawAutoCompleteData(oRs)


	inputID      = "Location"
   inputNm      = "Location"
   inputStyle   = "font-family:Tahoma;font-size:10px;width:200px;"
   DataSet      = "SC"
   DrawAutoCompleteBox()

	inputID      = "Location2"
	inputNm      = "Location2"
   inputStyle   = "font-family:Tahoma;font-size:10px;width:200px;"
   DataSet      = "SC"
   DrawAutoCompleteBox()


	response.write "<BR><BR>"


	inputID      = "Location3"
   inputNm      = "Location3"
   inputStyle   = "font-family:Tahoma;font-size:10px;width:200px;"
   DataSet      = "NC"
   DrawAutoCompleteBox()

	inputID      = "Location4"
	inputNm      = "Location4"
   inputStyle   = "font-family:Tahoma;font-size:10px;width:200px;"
   DataSet      = "NC"
   DrawAutoCompleteBox()
%>
</BODY>
</HTML>