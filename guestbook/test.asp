<!-- #include file="../common/general_includes.asp"-->
<%
	strCookies = Request.Cookies("EZGBAdmin")("name")
	
	If strCookies = "" Then

		Response.Redirect "admin_login.asp"
  
	End If

  If msg <> "" Then
    Call displayFancyMsg(getMessage(msg))
  End If
%>
<!-- #include file="../_includes/header.asp"-->
  <div id="main" class="container">
    <header>
      <h2>EZGuestbook</h2>
    </header>
    <ul>
      <li><a href="#admin"><span>Admin</span></a></li>
      <li><a href="admin_options.asp"><span>Manage options</span></a></li>
      <li><a href="#entries"><span>Manage entries</span></a></li>
      <li><a href="admin_chpwd.asp"><span>Change login info</span></a></li>
    </ul>
    <div id="entries" style="display:none;">
      <table>
        <tr>
          <td>
<%
Set Conn = Server.CreateObject("ADODB.Connection")
Call ConnOpen(Conn)	
Set rsCommon = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM tbl_guestbook ORDER BY guestbookID desc"
Call getTextRecordset(strSQL,rsCommon)
If Not rsCommon.EOF Then
If Request.Form("Page")="" then
PageNo = 1
Else
PageNo = Request.Form("Page")
End If
		
TotalRecs = rsCommon.RecordCount
' pagesize refers to the number of records displayed on each page
rsCommon.PageSize = 5
' the following line automatically figures out how many pages
' are needed for the recordset
TotalPages = CInt(rsCommon.PageCount)
' the next line tells the recordset which page you'd like to work with
rsCommon.AbsolutePage = PageNo
%>
            <div class="page_nav">
              <div class="page_nav_left" align="right">
                <% If PageNo > 1 then %>
                  <form method="post" action="admin.asp#entries">
                  <input type="hidden" name="Page" value="<%= PageNo-1 %>">
                  <input type="submit" value="<< Prev">
                  </form>
                <% Else %>
                  &nbsp;
                <% End If %>
              </div>
              <div class="page_nav_center" align="center">
                <span class="first">There are <%= TotalRecs %> entries! Page <%= PageNo %> Of <%= TotalPages %> pages.</span>
              </div>
              <div class="page_nav_right" align="left">
                <% If CInt(PageNo) < CInt(TotalPages) Then %>
                  <form method="post" action="admin.asp#entries">
                  <input type="hidden" name="Page" value="<%= PageNo+1 %>">
                  <input type="submit" value="Next >>">
                  </form>
                <% Else %>
                  &nbsp;
                <% End If %>  
              </div>
            </div>
<%
x = 0
For x = 1 to 5
		
banned = ""
Set rsCommon2 = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM tbl_IP_addr WHERE IP = '"&rsCommon("IP")&"'"
Call getTableRecordset("tbl_IP_addr",rsCommon2)
If Not rsCommon2.EOF Then
	banned = "<span style=""color:#ff0000"">IP Banned</span>"
End If
Call closeRecordset(rsCommon2)
			
			
%>        <div class="entry"><%
Set rsCommon2 = Server.CreateObject("ADODB.Recordset")
Call getTableRecordset("tbl_Fields",rsCommon2)
If Not rsCommon2.EOF Then
%>
              <div class="entry_block">
                <div class="entry_label"><strong>Date:</strong></div>
                <div class="entry_data"><%= rsCommon("gbdate") %></div>
              </div>		
<%
Do While Not rsCommon2.EOF
	If rsCommon2("field_show") = "yes" Then
		If rsCommon2("field_name") = "Site Visited" then
%>
              <div class="entry_block">
                <div class="entry_label"><strong>Site Visited:</strong></div>
                <div class="entry_data"><%= rsCommon("site") %>&nbsp;</div>
              </div>			
  <% ElseIf rsCommon2("field_name") = "Site Rating" then %>
              <div class="entry_block">
                <div class="entry_label"><strong>Site Rating:</strong></div>
                <div class="entry_data"><%= rsCommon("rate") %>&nbsp;</div>
              </div>
  <% ElseIf rsCommon2("field_name") = "Find Us?" then %>
              <div class="entry_block">
                <div class="entry_label"><strong>How did you find us?</strong></div>
                <div class="entry_data"><%= rsCommon("find") %>&nbsp;</div>
              </div>
  <% ElseIf rsCommon2("field_name") = "Comments" then %>
              <div class="entry_block">
                <div class="entry_label"><strong>Comments:</strong></div>
                <div class="entry_data"><%= rsCommon("comments") %>&nbsp;</div>
              </div>
  <% ElseIf rsCommon2("field_name") = "Name" then %>
              <div class="entry_block">
                <div class="entry_label"><strong>Name:</strong></div>
                <div class="entry_data"><%= rsCommon("name") %>&nbsp;</div>
              </div>
  <% ElseIf rsCommon2("field_name") = "Email" then %>
              <div class="entry_block">
                <div class="entry_label"><strong>Email:</strong></div>
                <div class="entry_data"><a href="mailto:<%= rsCommon("email") %>"><%= rsCommon("email") %></a>&nbsp;</div>
              </div>
  <% ElseIf rsCommon2("field_name") = "Website" then %>
              <div class="entry_block">
                <div class="entry_label"><strong>Website:</strong></div>
                <div class="entry_data"><a href="<%= rsCommon("website") %>"><%= rsCommon("website") %></a>&nbsp;</div>
              </div>
  <% ElseIf rsCommon2("field_name") = "AIM" then %>
              <div class="entry_block">
                <div class="entry_label"><strong>AIM:</strong></div>
                <div class="entry_data"><%= rsCommon("aim") %>&nbsp;</div>
              </div>				
  <% ElseIf rsCommon2("field_name") = "MSN" then %>
              <div class="entry_block">
                <div class="entry_label"><strong>MSN:</strong></div>
                <div class="entry_data"><%= rsCommon("msn") %>&nbsp;</div>
              </div>
  <% ElseIf rsCommon2("field_name") = "ICQ" then %>
              <div class="entry_block">
                <div class="entry_label"><strong>ICQ:</strong></div>
                <div class="entry_data"><%= rsCommon("icq") %>&nbsp;</div>
              </div>
  <% ElseIf rsCommon2("field_name") = "Age" then %>
              <div class="entry_block">
                <div class="entry_label"><strong>Age:</strong></div>
                <div class="entry_data"><%= rsCommon("age") %>&nbsp;</div>
              </div>				
  <% ElseIf rsCommon2("field_name") = "Location" then %>
              <div class="entry_block">
                <div class="entry_label"><strong>Location:</strong></div>
                <div class="entry_data"><%= rsCommon("loc") %>&nbsp;</div>
              </div>											
<%	
		End If
	End If
						
	rsCommon2.MoveNext
	If rsCommon2.EOF Then Exit Do
Loop
End If
Call closeRecordset(rsCommon2)
%>
              <div class="entry_block">
                <div class="entry_label"><strong>IP:</strong></div>
                <div class="entry_data"><%= rsCommon("IP") %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%= banned %></div>
              </div>
            </div>
            <div class="entry_links">
              <div class="entry_button" align="right">
                <form class="delentry" method="post">
                <input type="hidden" name="id" value="<%= rsCommon("guestbookID") %>">
                <input type="submit" name="submit" value="Delete this entry">
                </form>
              </div>
              <div class="entry_button" align="center">
                <form class="editentry" method="post">
                <input type="hidden" name="id" value="<%= rsCommon("guestbookID") %>">
                <input type="submit" name="submit" value="Edit this entry">
                </form>
              </div>
              <div class="entry_button" align="left">
                <form class="banip" method="post">
                <input type="hidden" name="id" value="<%= rsCommon("guestbookID") %>">
                <input type="submit" name="submit" value="Ban this IP">
                </form>
              </div>
            </div>
            <div class="spacer"></div>
<%
rsCommon.MoveNext
If rsCommon.EOF Then Exit for
Next
		
If Request.Form("Page")="" then
PageNo = 1
Else
PageNo = Request.Form("Page")
End If
		
TotalRecs = rsCommon.RecordCount
' pagesize refers to the number of records displayed on each page
rsCommon.PageSize = 5
' the following line automatically figures out how many pages
' are needed for the recordset
TotalPages = CInt(rsCommon.PageCount)
' the next line tells the recordset which page you'd like to work with
rsCommon.AbsolutePage = PageNo
%>	
            <div class="page_nav">
              <div class="page_nav_left" align="right">
                <% If PageNo > 1 then %>
                  <form method="post" action="admin.asp#ui-tabs-2">
                  <input type="hidden" name="Page" value="<%= PageNo-1 %>">
                  <input type="submit" value="<< Prev">
                  </form>
                <% Else %>
                  &nbsp;
                <% End If %>
              </div>
              <div class="page_nav_center" align="center">
                <span class="first">There are <%= TotalRecs %> entries! Page <%= PageNo %> Of <%= TotalPages %> pages.</span>
              </div>
              <div class="page_nav_right" align="left">
                <% If CInt(PageNo) < CInt(TotalPages) Then %>
                  <form method="post" action="admin.asp#ui-tabs-2">
                  <input type="hidden" name="Page" value="<%= PageNo+1 %>">
                  <input type="submit" value="Next >>">
                  </form>
                <% Else %>
                  &nbsp;
                <% End If %>  
              </div>
            </div>
<%
Else
Response.Write "<br /><span class=""first"">No Entries!</span>"
End If
	
Call closeRecordset(rsCommon)
Call ConnClose(Conn)
%> 
                          
        </td>
        </tr>
      </table>
    </div> 

    <% Call htmljunctionnews %>
    <br /><br />
 
  </div>
<!-- #include file="../_includes/footer.asp"-->