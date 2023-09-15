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

	intGuestbookID = 0
	If Trim(Request.QueryString("gid")) <> "" Then intGuestbookID = Trim(Request.QueryString("gid"))

	Set Conn = Server.CreateObject("ADODB.Connection")
	Call ConnOpen(Conn)

	If Request.Form("edit") = "yes" Then

		strName = DBEncode(Request.Form("name"))
		strEmail = DBEncode(Request.Form("email"))
		strWebsite = DBEncode(Request.Form("website"))
		strFaceBook = DBEncode(Request.Form("facebook"))
		strTwitter = DBEncode(Request.Form("twitter"))
		strAge = DBEncode(Request.Form("age"))
		strLocation = DBEncode(Request.Form("loc"))
		strSite = DBEncode(Request.Form("site"))
		strFind = DBEncode(Request.Form("find"))
		strRate = DBEncode(Request.Form("rate"))
		strComments = DBEncode(Request.Form("comments"))

		strSQL = "UPDATE "&msdbprefix&"guestbook SET name = '"&strName&"', email = '"&strEmail&"', website = '"&strWebsite&"', facebook = '"&strFaceBook&"', twitter = '"&strTwitter&"', age = '"&strAge&"', loc = '"&strLocation&"', site = '"&strSite&"', find = '"&strFind&"', rate = '"&strRate&"', comments = '"&strComments&"'  WHERE guestbookID = "&intGuestbookID
		Call getExecuteQuery(strSQL)

	    Request.Cookies("msg") = "ened"
		Response.Redirect "admin_entries.asp"

	End If

	Set rsCommon = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT * FROM "&msdbprefix&"guestbook WHERE guestbookID = "&intGuestbookID

	Call getTextRecordset(strSQL,rsCommon)
	If Not rsCommon.EOF Then
		strName = DBDecode(rsCommon("name"))
		strEmail = DBDecode(rsCommon("email"))
		strWebsite = DBDecode(rsCommon("website"))
		strFacebook = DBDecode(rsCommon("facebook"))
		strTwitter = DBDecode(rsCommon("twitter"))
		strAge = DBDecode(rsCommon("age"))
		strLocation = DBDecode(rsCommon("loc"))
		strSite = DBDecode(rsCommon("site"))
		strFind = DBDecode(rsCommon("find"))
		intRate = CInt(rsCommon("rate"))
		strComments = DBDecode(rsCommon("comments"))
		datDate = rsCommon("gbdate")
	End If
	Call closeRecordset(rsCommon)
%>
<!-- #include file="../includes/header.asp"-->
<div id="main" class="container">
  <header><h2>Edit Entry</h2></header>
<%
		Set rsCommon = Server.CreateObject("ADODB.Recordset")
		Call getTableRecordset(msdbprefix&"fields",rsCommon)
		If Not rsCommon.EOF Then
%>
	<form action="admin_edit.asp?gid=<%= intGuestbookID %>" method="post">
	<input type="hidden" name="edit" value="yes" />
  <div class="row">
    <div class="-3u 6u$ 12u(medium)" style="padding-bottom:10px;">
			<span><strong>Date:</strong> <%= datDate %></span>
    </div>
<%
			Do While Not rsCommon.EOF
				If rsCommon("field_show") = "yes" Then
			    If rsCommon("field_name") = "Site Visited" then
%>
		<div class="-3u 6u$ 12u(medium)" style="padding-bottom:10px;">
      <label for="site" style="margin-bottom:-3px;">Site Visited:</label>
      <div class="select-wrapper">
			  <% selectSite(strSite) %>
      </div>
		</div>			
<%		
			    ElseIf rsCommon("field_name") = "Site Rating" then
%>
	  <div class="-3u 6u$ 12u(medium)" style="padding-bottom:10px;">
      <label for="rate" style="margin-bottom:-3px;">Rate My Site</label>
      <div class="select-wrapper">
			  <% Call selectRate(intRate) %>
		  </div>
	  </div>
<%			
					ElseIf rsCommon("field_name") = "Find Us?" then
%>
		<div class="-3u 6u$ 12u(medium)" style="padding-bottom:10px;">
      <label for="find" style="margin-bottom:-3px;">How did you find us?</label>
      <div class="select-wrapper">
				<% selectFind(strFind) %>
			</div>
		</div>
	<%			
					ElseIf rsCommon("field_name") = "Comments" then
	%>
		<div class="-3u 6u$ 12u(medium)" style="padding-bottom:10px;">
			<label for="comments" style="margin-bottom:-3px;">Comments:</label>
			<textarea id="comments" name="comments" cols="30" rows="5"><%= strComments %></textarea>
		</div>
	<%			
					ElseIf rsCommon("field_name") = "Name" then
	%>
		<div class="-3u 6u$ 12u(medium)" style="padding-bottom:10px;">
			<label for="name" style="margin-bottom:-3px;">Name:</label>
			<input id="name" name="name" type="text" value="<%= strName %>" />
		</div>
	<%			
					ElseIf rsCommon("field_name") = "Email" then
	%>
		<div class="-3u 6u$ 12u(medium)" style="padding-bottom:10px;">
			<label for="email" style="margin-bottom:-3px;">Email:</label>
			<input id="email" name="email" type="text" value="<%= strEmail %>" />
		</div>
	<%			
					ElseIf rsCommon("field_name") = "Website" then
	%>
		<div class="-3u 6u$ 12u(medium)" style="padding-bottom:10px;">
			<label for="website" style="margin-bottom:-3px;">Website:</label>
			<input id="website" name="website" type="text" value="<%= strWebsite %>" />
		</div>
	<%			
					ElseIf rsCommon("field_name") = "facebook" then
	%>
		<div class="-3u 6u$ 12u(medium)" style="padding-bottom:10px;">
			<label for="facebook" style="margin-bottom:-3px;">Facebook:</label>
			<input id="facebook" name="facebook" type="text" value="<%= strFacebook %>" />
		</div>				
	<%			
					ElseIf rsCommon("field_name") = "twitter" then
	%>
		<div class="-3u 6u$ 12u(medium)" style="padding-bottom:10px;">
			<label for="twitter" style="margin-bottom:-3px;">Twitter:</label>
			<input id="twitter" name="twitter" type="text" value="<%= strTwitter %>" />
		</div>
	<%			
					ElseIf rsCommon("field_name") = "Age" then
	%>
		<div class="-3u 6u$ 12u(medium)" style="padding-bottom:10px;">
			<label for="age" style="margin-bottom:-3px;">Age:</label>
			<input id="age" name="age" type="text" value="<%= strAge %>" />
		</div>				
	<%			
					ElseIf rsCommon("field_name") = "Location" then
	%>
		<div class="-3u 6u$ 12u(medium)" style="padding-bottom:10px;">
			<label for="loc" style="margin-bottom:-3px;">Location:</label>
			<input id="loc" name="loc" type="text" value="<%= strLocation %>" />
		</div>											
	<%
					End If
				End If
						
				rsCommon.MoveNext
				If rsCommon.EOF Then Exit Do
			Loop
		End If
		Call closeRecordset(rsCommon)
    Call ConnClose(Conn)
	%>
		<div class="-3u 6u$ 12u(medium)">
			<input class="button" type="submit" value="Edit Entry" />
		</div>
  </div>
	</form>
</div>
<!-- #include file="../includes/footer.asp"-->