<!-- #include file="../includes/general_includes.asp"-->
<%
    strCookies = Request.Cookies("EZGBAdmin")("name")
	
    If strCookies = "" Then
        Response.Redirect "admin_login.asp"
    End If

    strIP = ""
    blnBanned = False
    intGuestookID = 0
    If Trim(Request.QueryString("gid")) <> "" Then intGuestookID = checkint(Trim(Request.QueryString("gid")))

    Set Conn = Server.CreateObject("ADODB.Connection")
    Call ConnOpen(Conn)

    Set rsCommon = Server.CreateObject("ADODB.Recordset")
    strSQL = "SELECT IP FROM "&msdbprefix&"guestbook WHERE guestbookID = "&intGuestookID

    Call getTextRecordset(strSQL,rsCommon)
    If Not rsCommon.EOF Then
        strIP = rsCommon("IP")
    End If
    Call closeRecordset(rsCommon)

    Set rsCommon = Server.CreateObject("ADODB.Recordset")
    strSQL = "SELECT * FROM "&msdbprefix&"IP WHERE IP = '"&strIP&"'"

    Call getTextRecordset(strSQL,rsCommon)
    If Not rsCommon.EOF Then
        blnBanned = True
    End If
    Call closeRecordset(rsCommon)

    If blnBanned = True Then

        strSQL = "DELETE FROM "&msdbprefix&"IP WHERE IP = '"&strIP&"'"
        Call getExecuteQuery(strSQL)

        Request.Cookies("msg") = "unban"
        Response.Redirect "admin_entries.asp"

    Else

        strSQL = "INSERT INTO "&msdbprefix&"IP(IP) values('"&strIP&"')"
        Call getExecuteQuery(strSQL)

        Request.Cookies("msg") = "ban"
        Response.Redirect "admin_entries.asp"

    End If

    Call ConnClose(Conn)

    Response.Redirect strRedirect
%>


