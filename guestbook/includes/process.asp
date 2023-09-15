<!-- #include file="general_includes.asp"-->
<%
on error resume next
	Set Conn = Server.CreateObject("ADODB.Connection")
	Call ConnOpen(Conn)

	If Request.Form("sign") = "yes" Then

		strName = DBEncode(Trim(Request.Form("name")))
		strEmail = Trim(Request.Form("email"))
		strWebsite = DBEncode(Trim(Request.Form("website")))
		strFacebook = DBEncode(Trim(Request.Form("facebook")))
		strTwitter = DBEncode(Trim(Request.Form("twitter")))
		strGooglePlus = DBEncode(Trim(Request.Form("googleplus")))
		strAge = checkint(Trim(Request.Form("age")))
		strLoc = DBEncode(Trim(Request.Form("location")))
		strSite = Trim(Request.Form("site"))
		strRate = Trim(Request.Form("rate"))
		strFind = Trim(Request.Form("find"))

	    strComments = RemoveHTML(Trim(Request.Form("comments")))
		strComments = DBEncode(strComments)

		strIP = Trim(Request.Form("IP"))

		If Left(strWebsite,7) = "http://" Then
		    strWebsite = Replace(strWebsite,"http://","")
		End If

		If Left(strWebsite,8) = "https://" Then
		    strWebsite = Replace(strWebsite,"https://","")
        End If

		Set rsCommon = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT IP FROM "&msdbprefix&"IP WHERE IP = '"&strIP&"'"
		Call getTextRecordset(strSQL,rsCommon)

		If Not rsCommon.EOF Then
			msg = "banip"
			Response.Redirect "../sign.asp?msg="&msg
		Else
			strSQL = "INSERT INTO "&msdbprefix&"guestbook ([name],[email],[website],[facebook],[twitter],[age],[loc],[site],[rate],[find],[IP],[BanIP],[gbdate],[comments]) VALUES ('"&strName&"','"&strEmail&"','"&strWebsite&"','"&strFacebook&"','"&strTwitter&"','"&strAge&"','"&strLoc&"','"& strSite&"','"&strRate&"','"&strFind&"','"&strIP&"','no','"&now&"','"&strComments&"')"
			Call getExecuteQuery(strSQL)

			Response.Redirect "../view.asp?msg=sus"
		End If
		Call closeRecordset(rsCommon)

  End If

  Call ConnClose(Conn)
%>