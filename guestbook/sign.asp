<!-- #include file="includes/general_includes.asp"-->
<!DOCTYPE HTML>
<html>
<head>
    <title>EZGuestbook - Sign</title>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
    <link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
    <link type="text/css" rel="stylesheet" href="assets/css/jquery.fancybox.css" />
    <link type="text/css" rel="stylesheet" href="assets/css/main.css" />
</head>

<body>
    <%
  If msg <> "" Then
    Call displayFancyMsg(getMessage(msg))
  End If

  Set Conn = Server.CreateObject("ADODB.Connection")
	Call ConnOpen(Conn)

	Set rsCommon = Server.CreateObject("ADODB.Recordset")
	Call getTableRecordset(msdbprefix&"options",rsCommon)
	If Not rsCommon.EOF Then
	 lngComCount = rsCommon("com_count")
	End If
	Call closeRecordset(rsCommon)

    %>
    <div id="main" class="container" style="margin-top: -50px;">
        <header style="text-align: center">
            <h2><%= strSiteTitle %> Guestbook</h2>
        </header>
        <div class="row">
            <div class="12u 12u(medium)" style="text-align: center; padding-bottom: 10px;">
                <a class="button" href="view.asp"><i class="fa fa-search"></i>View Guestbook</a>
            </div>
        </div>
        <div class="row 50%">
            <div class="12u 12u$(medium)">
                <form action="includes/process.asp" method="post">
                    <input type="hidden" name="sign" value="yes" />
                    <input type="hidden" name="IP" value="<%= Request.ServerVariables("REMOTE_HOST") %>" />
                    <div class="row">
                        <div class="-3u 6u 12u$(medium)">
                            <%
	Set rsCommon = Server.CreateObject("ADODB.Recordset")
	Call getTableRecordset(msdbprefix&"fields",rsCommon)
	If Not rsCommon.EOF Then
	  Do While Not rsCommon.EOF
		  If rsCommon("field_show") = "yes" Then
	      If rsCommon("field_name") = "Site Visited" then
                            %>
                            <div class="12u 12u$(medium)" style="padding-bottom: 30px;">
                                <div class="select-wrapper">
                                    <% Call selectSite("") %>
                                </div>
                            </div>
                            <%
			  ElseIf rsCommon("field_name") = "Site Rating" then
                            %>
                            <div class="12u 12u$(medium)" style="padding-bottom: 30px;">
                                <div class="select-wrapper">
                                    <% selectRate(0) %>
                                </div>
                                Please rate our site from 1 to 10, 10 being the best.
                            </div>
                            <%
			  ElseIf rsCommon("field_name") = "Find Us?" then
                            %>
                            <div class="12u 12u$(medium)" style="padding-bottom: 30px;">
                                <div class="select-wrapper">
                                    <% selectFind("") %>
                                </div>
                            </div>
                            <%
			  ElseIf rsCommon("field_name") = "Comments" then
                            %>
                            <div class="12u 12u$(medium)" style="padding-bottom: 30px;">
                                <textarea id="comments" name="comments" cols="30" rows="5" wrap="soft" onkeydown="textCounter(this.form.comments,this.form.remLen,<%= lngComCount %>);" onkeyup="textCounter(this.form.comments,this.form.remLen,<%= lngComCount %>);" placeholder="Comments:"></textarea>
                                <input type="text" id="remLen" name="remLen" size="3" maxlength="3" value="<%= lngComCount %>" readonly>
                                <label for="remLen">Characters left - HTML is not allowed!</label>
                            </div>
                            <%
			  ElseIf rsCommon("field_name") = "Website" Then
                            %>
                            <div class="12u 12u$(medium)" style="padding-bottom: 30px;">
                                <input id="website" name="<%= LCase(rsCommon("field_name")) %>" placeholder="Website: Example: www.htmljunction.com or aspjunction.com" type="text" />
                            </div>
                            <%
			  Else
                            %>
                            <div class="12u 12u$(medium)" style="padding-bottom: 30px;">
                                <input id="<%= LCase(rsCommon("field_name")) %>" placeholder="<%= rsCommon("field_name") %>:" name="<%= LCase(rsCommon("field_name")) %>" type="text" />
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
                            <div class="-4u 4u$ 12u$(medium)">
                                <input type="submit" value="Sign Guest Book" />
                            </div>
                        </div>
                    </div>
                </form>
            </div>
        </div>
    </div>
    <!-- REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE LICENSE AGREEMENT -->
    <footer id="footer">
        <div class="copyright">
            Powered by <a href="http://www.aspjunction.com">EZGuestbook</a> Copyright &copy; 2003 - <%= Year(Date) %> | <a href="http://<%= strDomain %>"><%= strSiteTitle %></a>
        </div>
    </footer>
    <!-- REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE LICENSE AGREEMENT -->
    <script type="text/javascript" src="https://code.jquery.com/jquery-1.12.4.js"></script>
    <script type="text/javascript" src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
    <script type="text/javascript" src="../assets/js/jquery.fancybox.js"></script>
    <script type="text/javascript" src="../assets/js/skel.min.js"></script>
    <script type="text/javascript" src="../assets/js/main.js"></script>
    <script language="javascript" type="text/javascript">
        $(document).ready(function () {
            $(".iframe").fancybox();
            $(".picimg").fancybox();
            $("#textmsg").fancybox();
            $("#textmsg").trigger('click');
        });

        function textCounter(field, countfield, maxlimit) {
            if (field.value.length > maxlimit) // if too long...trim it!
                field.value = field.value.substring(0, maxlimit);
            else // otherwise, update 'characters left' counter
                countfield.value = maxlimit - field.value.length;
        }
    </script>
</body>
</html>