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
%>
<!-- #include file="../includes/header.asp"-->
<div id="main" class="container">
    <div class="row">
        <div class="-1u 3u 12u$(medium)">
            <header>
                <h2>Manage Options</h2>
            </header>
        </div>
    </div>
    <div class="row">
        <div class="-1u 3u 12u$(medium)">
            <div class="box">
                <header>
                    <h3>Site Visited</h3>
                </header>
<%
        Set Conn = Server.CreateObject("ADODB.Connection")
        Call ConnOpen(Conn)

        Set rsCommon = Server.CreateObject("ADODB.Recordset")
        Call getTableRecordset(msdbprefix&"site",rsCommon)
        If Not rsCommon.EOF Then
            Response.Write "<form action=""admin_process.asp"" method=""post"">"&vbcrlf
            Response.Write "<div class=""row"">"&vbcrlf
            Do While Not rsCommon.EOF

                intSiteID = 0
                strSiteName = ""
                strSiteNameCleaned = ""

                intSiteID = rsCommon("siteID")
                strSiteName = DBDecode(rsCommon("site_name"))
                strSiteNameCleaned = Replace(strSiteName,"?","")
                strSiteNameCleaned = Replace(strSiteNameCleaned," ","")
                strSiteNameCleaned = LCase(strSiteNameCleaned)

                Response.Write "  <div class=""12u 12u(medium)"">"&vbcrlf
                Response.Write "    <input type=""checkbox"" id="""&strSiteNameCleaned&""" name=""del"" value="""&intSiteID&""" />"&vbcrlf
                Response.Write "    <label for="""&strSiteNameCleaned&""">"&strSiteName&"</label>"&vbcrlf
                Response.Write "  </div>"&vbcrlf

                rsCommon.MoveNext
                If rsCommon.EOF Then Exit Do
            Loop

            Response.Write "  <div class=""12u$ 12u(medium)"">"&vbcrlf
            Response.Write "    <input class=""button fit"" type=""submit"" name=""delsite"" value=""Delete Selected"">"&vbcrlf
            Response.Write "  </div>"&vbcrlf
            Response.Write "</div>"&vbcrlf
            Response.Write "</form>"&vbcrlf

        End If
        Call closeRecordset(rsCommon)

        Response.Write "<form action=""admin_process.asp"" method=""post"">"&vbcrlf
        Response.Write "<div class=""row"">"&vbcrlf
        Response.Write "	<div class=""12u 12u(medium)"" style=""padding-bottom:10px;"">"&vbcrlf
        Response.Write "    <input type=""text"" id=""newoption"" name=""newoption"" size=""20"">"&vbcrlf
        Response.Write "  </div>"&vbcrlf
        Response.Write "  <div class=""12u 12u(medium)"" style=""text-align:center;"">"&vbcrlf
        Response.Write "    <input class=""button fit"" type=""submit"" name=""addsiteoption"" value=""Add New Option"">"&vbcrlf
        Response.Write "  </div>"&vbcrlf
        Response.Write "</div>"&vbcrlf
        Response.Write "</form>"&vbcrlf


        Set rsCommon = Server.CreateObject("ADODB.Recordset")
        Call getTableRecordset(msdbprefix&"options",rsCommon)
        If Not rsCommon.EOF Then
            If rsCommon("orderby") = "desc" Then
                strChecked = "checked"
            Else
                strChecked = ""
            End If

            intEntriesPage = rsCommon("entries_page")
            lngComCount = rsCommon("com_count")
        End If
%>
            </div>
            <div class="box">
                <header>
                    <h4>More Options</h4>
                </header>
                <form action="admin_process.asp" method="post">
                    <div class="row">
                        <div class="12u$">
                            <input type="checkbox" id="orderby" name="orderby" <%= strChecked %> />
                            <label for="orderby">Show newest entry first?</label>
                        </div>
                        <div class="12u$">
                            <input type="text" id="entriespage" name="entriespage" value="<%= intEntriesPage %>" size="3" />
                            <label for="entriespage">Number of entries to display per page</label>
                        </div>
                        <div class="12u$">
                            <input type="text" id="comcount" name="comcount" value="<%= lngComCount %>" size="3" />
                            <label for="comcount">Number of characters to allow in comment box</label>
                        </div>
                        <div class="12u$">
                            <input class="button fit" type="submit" name="moreoptions" value="Submit" />
                        </div>
                    </div>
                </form>
            </div>
        </div>
        <div class="3u 12u$(medium)">
            <div class="box">
                <header>
                    <h4>How did you find us?</h4>
                </header>
<%
            Set rsCommon = Server.CreateObject("ADODB.Recordset")
            Call getTableRecordset(msdbprefix&"find",rsCommon)
            If Not rsCommon.EOF Then
                Response.Write "<form action=""admin_process.asp"" method=""post"">"&vbcrlf
                Response.Write "<input type=""hidden"" name=""find"" value=""yes"">"&vbcrlf
                Response.Write "<div class=""row"">"&vbcrlf
                Do While Not rsCommon.EOF

                    Response.Write "  <div class=""12u$ 12u(medium)"">"&vbcrlf
                    Response.Write "    <input type=""checkbox"" id="""&DBDecode(rsCommon("find_name"))&""" name=""del"" value="""&DBDecode(rsCommon("findID"))&""">"&vbcrlf
                    Response.Write "    <label for="""&rsCommon("find_name")&""">"&rsCommon("find_name")&"</label>"&vbcrlf
                    Response.Write "  </div>"&vbcrlf
                    rsCommon.MoveNext
                    If rsCommon.EOF Then Exit Do

                Loop
                Response.Write "  <div class=""12u$ 12u(medium)"">"&vbcrlf
                Response.Write "    <input class=""button fit"" type=""submit"" name=""delfind"" value=""Delete Selected"">"&vbcrlf
                Response.Write "  </div>"&vbcrlf
                Response.Write "</div>"&vbcrlf
                Response.Write "</form>"&vbcrlf
            End If
            Call closeRecordset(rsCommon)

            Response.Write "<form action=""admin_process.asp"" method=""post"">"&vbcrlf
            Response.Write "<div class=""row"">"&vbcrlf
            Response.Write "	<div class=""12u 12u(medium)"" style=""padding-bottom:10px;"">"&vbcrlf
            Response.Write "    <input type=""text"" id=""newoption"" name=""newoption"" size=""20"">"&vbcrlf
            Response.Write "  </div>"&vbcrlf
            Response.Write "  <div class=""12u 12u(medium)"" style=""text-align:center;"">"&vbcrlf
            Response.Write "    <input class=""button fit"" type=""submit"" name=""addfindoption"" value=""Add New Option"">"&vbcrlf
            Response.Write "  </div>"&vbcrlf
            Response.Write "</div>"&vbcrlf
            Response.Write "</form>"&vbcrlf
%>
            </div>
        </div>
        <div class="3u$ 12u$(medium)">
            <div class="box">
                <header>
                    <h4>Guestbook Fields</h4>
                </header>
<%
    Set rsCommon = Server.CreateObject("ADODB.Recordset")
    Call getTableRecordset(msdbprefix&"fields",rsCommon)
    If Not rsCommon.EOF Then
        Response.Write "<form action=""admin_process.asp"" method=""post"">"&vbcrlf
        Response.Write "<div class=""row"">"&vbcrlf
        Do While Not rsCommon.EOF
            strChecked = ""
            intFieldID = 0
            strFieldName = ""
            strFieldNameCleaned = ""

            intFieldID = rsCommon("fieldID")

            strFieldName = DBDecode(rsCommon("field_name"))
            strFieldNameCleaned = Replace(strFieldName,"?","")
            strFieldNameCleaned = Replace(strFieldNameCleaned," ","")
            strFieldNameCleaned = LCase(strFieldNameCleaned)

            If rsCommon("field_show") = "yes" Then strChecked = "checked"

            Response.Write "  <div class=""8u 12u$(medium)"">"&vbcrlf
            Response.Write "    <input type=""checkbox"" id="""&strFieldNameCleaned&""" name=""show"" value="""&intFieldID&""" "&strChecked&" >"&vbcrlf
            Response.Write "    <label for="""&strFieldNameCleaned&""">"&strFieldName&"</label>"&vbcrlf
            Response.Write "  </div>"&vbcrlf
            Response.Write "  <div class=""4u$ 12u$(medium)"">"&vbcrlf
            If rsCommon("field_show") = "yes" Then
                Response.Write "Showing"
            Else
                Response.Write "Hidden"
            End If
            Response.Write "  </div>"&vbcrlf
            rsCommon.MoveNext
            If rsCommon.EOF Then Exit Do
        Loop
        Response.Write "	<div class=""12u 12u(medium)"">"&vbcrlf
        Response.Write "    <input class=""button fit"" type=""submit"" name=""checkfields"" value=""Check Fields"">"&vbcrlf
        Response.Write "  </div>"&vbcrlf
        Response.Write "</div>"&vbcrlf
        Response.Write "</form>"&vbcrlf
    End If

    Call closerecordset(rsCommon)

    strChecked = ""
    Set rsCommon = Server.CreateObject("ADODB.Recordset")
    Call getTableRecordset(msdbprefix&"options",rsCommon)
    If NOT rsCommon.EOF Then
        strOrderby = rsCommon("orderby")
        intEntriesPage = rsCommon("entries_page")
        lngComCount = rsCommon("com_count")
    End If
    Call closerecordset(rsCommon)
    If strOrderby = "desc" Then strChecked = "checked"
%>
            </div>
        </div>
    </div>
</div>
<!-- #include file="../includes/footer.asp"-->