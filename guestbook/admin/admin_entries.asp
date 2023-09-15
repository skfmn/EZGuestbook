<!-- #include file="../includes/general_includes.asp"-->
<%
    strCookies = Request.Cookies("EZGBAdmin")("name")
	
    If strCookies = "" Then
        Response.Redirect "admin_login.asp"
    End If 
%>
<!-- #include file="../includes/header.asp"-->
<div id="main" class="container">
  <div class="row">
    <div class="-3u 6u$ 12u(medium)">
        <header>
            <h2>Manage Entries</h2>
        </header>
    </div>
  </div>
  <div class="row">
    <div class="-3u 6u 12u(medium)">
<%
    Set Conn = Server.CreateObject("ADODB.Connection")
    Call ConnOpen(Conn)
  	
    Set rsCommon = Server.CreateObject("ADODB.Recordset")
    strSQL = "SELECT * FROM "&msdbprefix&"guestbook ORDER BY guestbookID desc"

    Call getTextRecordset(strSQL,rsCommon)
    If Not rsCommon.EOF Then

	    If Request.QueryString("page") = "" then
		    PageNo = 1
	    Else
		    PageNo = Request.QueryString("page")
	    End If
		
	    TotalRecs = rsCommon.RecordCount
	    rsCommon.PageSize = 5
	    TotalPages = CInt(rsCommon.PageCount)
	    rsCommon.AbsolutePage = PageNo

	    x = 0
	    For x = 1 to 5
		
            blnBanned = False
            Set rsCommon2 = Server.CreateObject("ADODB.Recordset")
            strSQL = "SELECT * FROM "&msdbprefix&"IP WHERE IP = '"&rsCommon("IP")&"'"
            Call getTableRecordset(msdbprefix&"IP",rsCommon2)
            If Not rsCommon2.EOF Then
	            blnBanned = True
            End If
            Call closeRecordset(rsCommon2)	
%>
      <div class="table-wrapper">
			  <table class="alt" style="border:#dddddd solid 1px;">
          <tbody>      
<%
            Set rsCommon2 = Server.CreateObject("ADODB.Recordset")
            Call getTableRecordset(msdbprefix&"fields",rsCommon2)
            If Not rsCommon2.EOF Then
%>
            <tr>
              <td style="text-align:left;width:30%;"><strong>Date:</strong></td>
              <td style="text-align:left;width:70%;"><%= rsCommon("gbdate") %></td>
            </tr>		
<%
                Do While Not rsCommon2.EOF
	                If rsCommon2("field_show") = "yes" Then
		                If rsCommon2("field_name") = "Site Visited" then
%>
            <tr>
              <td style="text-align:left;width:30%;"><strong>Site Visited:</strong></td>
              <td style="text-align:left;width:70%;"><%= rsCommon("site") %>&nbsp;</td>
            </tr>			
<%          ElseIf rsCommon2("field_name") = "Site Rating" then %>
            <tr>
              <td style="text-align:left;width:30%;"><strong>Site Rating:</strong></td>
              <td style="text-align:left;width:70%;"><%= rsCommon("rate") %>&nbsp;</td>
            </tr>
<%          ElseIf rsCommon2("field_name") = "Find Us?" then %>
            <tr>
              <td style="text-align:left;width:30%;"><strong>How did you find us?</strong></td>
              <td style="text-align:left;width:70%;"><%= rsCommon("find") %>&nbsp;</td>
            </tr>
<%          ElseIf rsCommon2("field_name") = "Comments" then %>
            <tr>
              <td style="text-align:left;width:30%;"><strong>Comments:</strong></td>
              <td style="text-align:left;width:70%;"><%= rsCommon("comments") %>&nbsp;</td>
            </tr>
<%          ElseIf rsCommon2("field_name") = "Name" then %>
            <tr>
              <td style="text-align:left;width:30%;"><strong>Name:</strong></td>
              <td style="text-align:left;width:70%;"><%= rsCommon("name") %>&nbsp;</td>
            </tr>
<%          ElseIf rsCommon2("field_name") = "Email" then %>
            <tr>
              <td style="text-align:left;width:30%;"><strong>Email:</strong></td>
              <td style="text-align:left;width:70%;"><a href="mailto:<%= rsCommon("email") %>"><%= rsCommon("email") %></a>&nbsp;</td>
            </tr>
<%          ElseIf rsCommon2("field_name") = "Website" then %>
            <tr>
              <td style="text-align:left;width:30%;"><strong>Website:</strong></td>
              <td style="text-align:left;width:70%;"><a href="http://<%= rsCommon("website") %>"><%= rsCommon("website") %></a>&nbsp;</td>
            </tr>
<%          ElseIf rsCommon2("field_name") = "facebook" then %>
            <tr>
              <td style="text-align:left;width:30%;"><strong>Facebook:</strong></td>
              <td style="text-align:left;width:70%;"><%= rsCommon("facebook") %>&nbsp;</td>
            </tr>				
<%          ElseIf rsCommon2("field_name") = "twitter" then %>
            <tr>
              <td style="text-align:left;width:30%;"><strong>Twitter:</strong></td>
              <td style="text-align:left;width:70%;"><%= rsCommon("twitter") %>&nbsp;</td>
            </tr>
<%          ElseIf rsCommon2("field_name") = "googleplus" then %>
            <tr>
              <td style="text-align:left;width:30%;"><strong>G+:</strong></td>
              <td style="text-align:left;width:70%;"><%= rsCommon("googleplus") %>&nbsp;</td>
            </tr>
<%          ElseIf rsCommon2("field_name") = "Age" then %>
            <tr>
              <td style="text-align:left;width:30%;"><strong>Age:</strong></td>
              <td style="text-align:left;width:70%;"><%= rsCommon("age") %>&nbsp;</td>
            </tr>				
<%          ElseIf rsCommon2("field_name") = "Location" then %>
            <tr>
              <td style="text-align:left;width:30%;"><strong>Location:</strong></td>
              <td style="text-align:left;width:70%;"><%= rsCommon("loc") %>&nbsp;</td>
            </tr>											
<%	
		                End If
		            End If
						
		            rsCommon2.MoveNext
		            If rsCommon2.EOF Then Exit Do
	            Loop
            End If
            Call closeRecordset(rsCommon2)
      
            banned = ""
            If blnBanned = True Then banned = "<span style=""color:#ff0000"">IP Banned</span>"
%>
            <tr>
              <td style="text-align:left;width:30%;"><strong>IP:</strong></td>
              <td style="text-align:left;width:70%;"><%= rsCommon("IP") %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%= banned %></td>
            </tr>
            <tr>
              <td colspan="2" style="margin:0px;padding:0px;border:0px;text-align:center;">
                <table style="margin:0px;padding:0px;border:0px;">
                  <tbody>
                    <tr>
                      <td>
                        <a class="button" onclick="return confirmSubmit('Are you sure you want to delete this entry?','admin_delete.asp?gid=<%= rsCommon("guestbookID") %>')"><i class="fa fa-times-circle"></i> Delete</a>
                      </td>
                      <td>
                        <a class="button" href="admin_edit.asp?gid=<%= rsCommon("guestbookID") %>"><i class="fa fa-edit"></i> Edit</a>
                      </td>
                      <td>
                        <% If blnBanned = True Then %>
                        <a class="button" href="admin_banip.asp?gid=<%= rsCommon("guestbookID") %>"><i class="fa fa-circle-o"></i> Un-ban</a>
                        <% Else %>
                        <a class="button" href="admin_banip.asp?gid=<%= rsCommon("guestbookID") %>"><i class="fa fa-ban"></i> Ban</a>
                        <% End If %>
                      </td>
                    </tr>
                  </tbody>
                </table> 
              </td>
            </tr>
          </tbody>
        </table>
      </div>
<%
			rsCommon.MoveNext
			If rsCommon.EOF Then Exit for
		Next
    Else
%>
        <table>
            <tr>
              <td>No Entries</td>
            </tr>	
         </table>
<%
	End If
    Call closeRecordset(rsCommon)
	Call ConnClose(Conn)
%>
    </div>
  </div>
</div>
<div id="main" class="container">
<%
    If PageNo <> 0 Then Call adminPagination(PageNo)
%>
</div>
<!-- #include file="../includes/footer.asp"-->