<%
	Dim Conn, ConnStr, rsCommon, rsCommon2, rsCommon3, strSQL, fso
	
	Dim strName, strEmail, strWebsite, strAimname, strMsnname, strIcqnum, banned, rsBan
	Dim strAge, strLoc, strSite, strRate, strFind, strComments, rsEdit, strSiteTitle, strDomainname
	Dim strFieldIDs, myItem, strSplit, PageNo, TotalRecs, TotalPages, x, strDelete, SplitDelete
	Dim strItem, strIP, rsSign, msg, strVersion, intEntriesPage, strOrderby, lngComCount

  Const ForReading = 1
  Const TristateUseDefault = -2

	
	msg = Trim(Request.QueryString("msg"))
	
	Response.ExpiresAbsolute = Now() - 2
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "No-Store"
	Response.AddHeader "If-Modified-Since",now
	Response.AddHeader "Last-Modified",now
	Response.Expires = 0

	If Request.ServerVariables("HTTPS") = "off" Then
	    strHTTP = "http://"
	Else
	    strHTTP = "https://"
	End If

%>
