<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
' *** Logout the current user.
Response.Cookies("EZGBAdmin").Expires = Date - 1
Session.Abandon()
Response.Redirect "admin_login.asp"
%>