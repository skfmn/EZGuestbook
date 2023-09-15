<!-- #include file="../includes/general_includes.asp"-->
<%
	strCookies = Request.Cookies("EZGBAdmin")("name")
	
	If strCookies = "" Then
	    Response.Redirect "admin_login.asp"
	End If

	intGuestookID = 0
	If Trim(Request.QueryString("gid")) <> "" Then intGuestookID = checkint(Trim(Request.QueryString("gid")))

	Set Conn = Server.CreateObject("ADODB.Connection")
	Call ConnOpen(Conn)

	strSQL = "DELETE FROM "&msdbprefix&"guestbook WHERE guestbookID = "&intGuestookID
	Call getExecuteQuery(strSQL)
	Call ConnClose(Conn)

	Request.Cookies("msg") = "ed"
	Response.Redirect "admin_entries.asp"
%>


