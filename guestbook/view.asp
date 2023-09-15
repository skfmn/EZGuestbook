<!-- #include file="includes/general_includes.asp"-->
<!DOCTYPE HTML>
<html>
<head>
    <title>EZGuestbook - View</title>
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
    %>
    <div id="main" class="container" style="margin-top: -50px;">
        <header style="text-align: center">
            <h2><%= strSiteTitle %> Guestbook</h2>
        </header>
        <div class="row">
            <div class="12u 12u(medium)" style="text-align: center; padding-bottom: 20px;">
                <a class="button" href="sign.asp"><i class="fa fa-pencil"></i>Sign Guestbook</a>
            </div>
        </div>
        <div class="row 50%">
            <%
on error resume next

    Set Conn = Server.CreateObject("ADODB.Connection")
    Call ConnOpen(Conn)

    Set rsCommon = Server.CreateObject("ADODB.Recordset")
    Call getTableRecordset(msdbprefix&"options",rsCommon)
    If NOT rsCommon.EOF Then
        strOrderby = rsCommon("orderby")
        intEntriesPage = rsCommon("entries_page")
    End If
    Call closeRecordset(rsCommon)

    Set rsCommon = Server.CreateObject("ADODB.Recordset")
    strSQL = "SELECT * FROM "&msdbprefix&"guestbook ORDER BY guestbookID "&strOrderby
    Call getTextRecordset(strSQL,rsCommon)
    If Not rsCommon.EOF Then

        If Request.QueryString("page") = "" then
	        PageNo = 1
        Else
	        PageNo = Request.QueryString("page")
        End If

        TotalRecs = rsCommon.RecordCount
        ' pagesize refers to the number of records displayed on each page
        rsCommon.PageSize = intEntriesPage
        ' the following line automatically figures out how many pages
        ' are needed for the recordset
        TotalPages = CInt(rsCommon.PageCount)
        ' the next line tells the recordset which page you'd like to work with
        rsCommon.AbsolutePage = PageNo

        If PageNo <> 0 Then Call pagination(PageNo)

        x = 0
        For x = 1 to Cint(intEntriesPage)
            %>
            <div class="-3u 6u$ 12u(medium)">
                <div class="table-wrapper">
                    <table class="alt" style="border: #dddddd solid 1px;">
                        <tbody>
                            <tr>
                                <td style="text-align: left; width: 30%;"><strong>Date:</strong></td>
                                <td style="text-align: left; width: 70%;"><%= rsCommon("gbdate") %></td>
                            </tr>
                            <%
            Set rsCommon2 = Server.CreateObject("ADODB.Recordset")
            Call getTableRecordset(msdbprefix&"fields",rsCommon2)
            If Not rsCommon2.EOF Then
	            Do While Not rsCommon2.EOF
		            If rsCommon2("field_show") = "yes" Then
	                If rsCommon2("field_name") = "Site Visited" then
                            %>
                            <tr>
                                <td style="text-align: left;"><strong>Site Visited:</strong></td>
                                <td style="text-align: left;"><%= rsCommon("site") %></td>
                            </tr>
                            <%
			      ElseIf rsCommon2("field_name") = "Site Rating" then
                            %>
                            <tr>
                                <td style="text-align: left;"><strong>Site Rating:</strong></td>
                                <td style="text-align: left;"><%= rsCommon("rate") %></td>
                            </tr>
                            <%
			      ElseIf rsCommon2("field_name") = "Find Us?" then
                            %>
                            <tr>
                                <td style="text-align: left;"><strong>How did you find us?</strong></td>
                                <td style="text-align: left;"><%= rsCommon("find") %></td>
                            </tr>
                            <%
			      ElseIf rsCommon2("field_name") = "Comments" then
                            %>
                            <tr>
                                <td style="text-align: left;"><strong>Comments:</strong></td>
                                <td style="text-align: left;"><%= rsCommon("comments") %></td>
                            </tr>
                            <%
			      ElseIf rsCommon2("field_name") = "Name" then
                            %>
                            <tr>
                                <td style="text-align: left;"><strong>Name:</strong></td>
                                <td style="text-align: left;"><%= rsCommon("name") %></td>
                            </tr>
                            <%
			      ElseIf rsCommon2("field_name") = "Email" then
                            %>
                            <tr>
                                <td style="text-align: left;"><strong>Email:</strong></td>
                                <td style="text-align: left;"><a href="mailto:<%= rsCommon("email") %>"><%= rsCommon("email") %></a></td>
                            </tr>
                            <%
			      ElseIf rsCommon2("field_name") = "Website" then
                            %>
                            <tr>
                                <td style="text-align: left;"><strong>Website:</strong></td>
                                <td style="text-align: left;"><a href="<%= rsCommon("website") %>"><%= rsCommon("website") %></a></td>
                            </tr>
                            <%
			      ElseIf rsCommon2("field_name") = "facebook" then
                            %>
                            <tr>
                                <td style="text-align: left;"><strong>Facebook:</strong></td>
                                <td style="text-align: left;"><a href="https://www.facebook.com/<%= rsCommon("facebook") %>"><%= rsCommon("facebook") %></a></td>
                            </tr>
                            <%
			      ElseIf rsCommon2("field_name") = "twitter" then
                            %>
                            <tr>
                                <td style="text-align: left;"><strong>Twitter:</strong> </td>
                                <td style="text-align: left;"><a href="https://twitter.com/<%= rsCommon("twitter") %>"><%= rsCommon("twitter") %></a></td>
                            </tr>
                            <%
			      ElseIf rsCommon2("field_name") = "googleplus" then
                            %>
                            <tr>
                                <td style="text-align: left;"><strong>Goggle+:</strong></td>
                                <td style="text-align: left;"><a href="https://plus.google.com/<%= rsCommon("googleplus") %>"><%= rsCommon("googleplus") %></a></td>
                            </tr>
                            <%
			      ElseIf rsCommon2("field_name") = "Age" then
                            %>
                            <tr>
                                <td style="text-align: left;"><strong>Age:</strong></td>
                                <td style="text-align: left;"><%= rsCommon("age") %></td>
                            </tr>
                            <%
			      ElseIf rsCommon2("field_name") = "Location" then
                            %>
                            <tr>
                                <td style="text-align: left;"><strong>Location:</strong></td>
                                <td style="text-align: left;"><%= rsCommon("loc") %></td>
                            </tr>
                            <%
			      End If
		      End If

	        rsCommon2.MoveNext
		      If rsCommon2.EOF Then Exit Do
	      Loop
      End If
      Call closeRecordset(rsCommon2)
                            %>
                        </tbody>
                    </table>
                </div>
            </div>

            <%
	        rsCommon.MoveNext
	        If rsCommon.EOF Then Exit for
        Next
    Else
%>
            <div class="-3u 6u 12u(medium)">
                <div class="table-wrapper">
                    <table>
                        <tr>
                          <td>No Entries</td>
                        </tr>	
                     </table>
                </div>
            </div>
<%
    End If
    Call closeRecordset(rsCommon)
    Call ConnClose(Conn)

    If PageNo <> 0 Then Call pagination(PageNo)
%>
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