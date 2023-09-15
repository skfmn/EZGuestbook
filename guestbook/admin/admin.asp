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

  Set fso = Server.CreateObject("Scripting.FileSystemObject")

	If fso.FolderExists(Server.MapPath("/guestbook/install")) Then
		fso.DeleteFolder(Server.MapPath("/guestbook/install"))
	End If

	Set fso = Nothing
%>
<!-- #include file="../includes/header.asp"-->
<div id="main" class="container">
    <header>
        <h1 style="text-align: center;font-size:30px">EZGuestbook</h1>
        <h4 style="text-align: center;">Choose an Option below</h4>
    </header>
    <div class="row">
         <div class="-3u 3u 12u$(medium)">
            <ul class="alt">
                <li>
                    <a class="button fit" href="admin_entries.asp">
                        <span>Manage entries</span>
                    </a>
                </li>
            </ul>
        </div>
        <div class="3u$ 12u$(medium)">
            <ul class="alt">
                <li>
                    <a class="button fit" href="admin_options.asp">
                        <span>Manage options</span>
                    </a>
                </li>
            </ul>
        </div>
        <div class="-3u 3u 12u$(medium)">
            <ul class="alt">
                <li>
                    <a class="button fit" href="admin_settings.asp">
                        <span>Manage Settings</span>
                    </a>
                </li>
            </ul>
        </div>
        <div class="3u$ 12u$(medium)">
            <ul class="alt">
                <li>
                    <a class="button fit" href="admin_manage.asp">
                        <span>Manage Admins</span>
                    </a>
                </li>
            </ul>
        </div>
        <div class="-3u 6u 12u$(medium)">
            <%= getResponse("http://www.aspjunction.com/gnews.asp?gbv="& strVersion&"") %>
        </div>

    </div>
</div>
<!-- #include file="../includes/footer.asp"-->