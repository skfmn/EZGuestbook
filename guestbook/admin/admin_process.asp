<!-- #include file="../includes/general_includes.asp"-->
<%
	strCookies = Request.Cookies("EZGBAdmin")("name")
	
	If strCookies = "" Then
	    Response.Redirect "admin_login.asp"
	End If

    msg = ""
    msg = Trim(Request.Cookies("msg"))

	If msg <> "" Then
		Call displayFancyMsg(getMessage(msg))
        Response.Cookies("msg") = ""
	End If

	Set Conn = Server.CreateObject("ADODB.Connection")
	Call ConnOpen(Conn)
	
	If Trim(Request.Form("checkfields")) <> "" Then
	
		strFieldIDs = Trim(Request.Form("show"))

		Set rsCommon = Server.CreateObject("ADODB.Recordset")
		Call getTableRecordset(msdbprefix&"fields",rsCommon)
		If Not rsCommon.EOF Then
			Do While Not rsCommon.EOF

				rsCommon("field_show") = "no"
				rsCommon.Update

				strSplit = Split(strFieldIDs,",")
				For Each myItem in strSplit

					intFieldID = rsCommon("fieldID")
					If CInt(intFieldID)  = CInt(myItem) Then
					    rsCommon("field_show") = "yes"
					    rsCommon.Update
					End If

				Next

				strFieldIDs = Join(strSplit,",")
				rsCommon.MoveNext
				If rsCommon.EOF Then Exit Do
			Loop
		End If
		Call closeRecordset(rsCommon)

	    Response.Cookies("msg") = "fch"
		Response.Redirect "admin_options.asp"

	End If
	
	If Trim(Request.Form("moreoptions")) <> "" Then
	
		intEntriesPage = Trim(Request.Form("entriespage"))
		lngComCount = Trim(Request.Form("comcount"))

		If Trim(Request.Form("orderby")) = "on" Then
		    strOrderby = "desc"
		Else
		    strOrderby = "asc"
		End If
	  
		Set rsCommon = Server.CreateObject("ADODB.Recordset")
		Call getTableRecordset(msdbprefix&"options",rsCommon)
		If NOT rsCommon.EOF Then
		    rsCommon("orderby") = strOrderby
		    rsCommon("entries_page") = intEntriesPage
		    rsCommon("com_count") = lngComCount
		    rsCommon.Update
		End If
		Call closeRecordset(rsCommon)
		
	    Response.Cookies("msg") = "opa"
		Response.Redirect "admin_options.asp" 

	End If
	
	If Trim(Request.Form("addsiteoption")) <> "" Then
        Call getExecuteQuery("INSERT INTO "&msdbprefix&"site([site_name]) values('"&DBEncode(Request.Form("newoption"))&"')")
	    Response.Cookies("msg") = "opa"
		Response.Redirect "admin_options.asp"
	End If
	
	If Trim(Request.Form("addfindoption")) <> "" Then
	    Call getExecuteQuery("INSERT INTO "&msdbprefix&"find([find_name]) values('"&DBEncode(Request.Form("newoption"))&"')")
	    Response.Cookies("msg") = "opa"
		Response.Redirect "admin_options.asp"
	End If
  
	If Request.Form("delsite") <> "" Then
		strDelete = Request.Form("Del")
		SplitDelete = Split(strDelete,",")
		
		For Each strItem in SplitDelete
		    strSQL = "DELETE FROM "&msdbprefix&"site WHERE siteID = "&strItem
		    Call getExecuteQuery(strSQL)
		Next
		
		strDelete = Join(SplitDelete,",")

	    Response.Cookies("msg") = "dls"
		Response.Redirect "admin_options.asp"

	End If
	
	If Request.Form("delfind") <> "" Then

		strDelete = Request.Form("del")
		SplitDelete = Split(strDelete,",")
	
		For Each strItem in SplitDelete
		    strSQL = "DELETE FROM "&msdbprefix&"find WHERE findID = "&checkint(strItem)
		    Call getExecuteQuery(strSQL)
		Next
	
		strDelete = Join(SplitDelete,",")

	    Response.Cookies("msg") = "dlf"
		Response.Redirect "admin_options.asp"

	End If
  
	Call ConnClose(Conn)
%>

