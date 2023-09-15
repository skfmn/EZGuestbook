<!-- #include file="../includes/general_includes.asp"-->
<!-- #include file="../includes/hash.asp"-->
<%

  strUsername = Trim(Request.Form("name"))
  strPassword = Trim(Request.Form("pwd"))

  If strUsername <> "" Then

    Set Conn = Server.CreateObject("ADODB.Connection")
    Set rsCommon = Server.CreateObject("ADODB.Recordset")

    Conn.Open ConnStr

    strSQL = "SELECT salt FROM "&msdbprefix&"admin WHERE name = '"&strUsername&"'"
    rsCommon.Open strSQL, Conn, 3, 3, &H0001
	  If Not rsCommon.EOF Then
		  strSalt = rsCommon("salt")
	  Else
		  Response.Redirect "admin_login.asp"
	  End If 
	  rsCommon.Close: Set rsCommon = Nothing

	  strEncrPassword = HashEncode(strPassword&strSalt)

    Set rsCommon = Server.CreateObject("ADODB.Recordset") 
    strSQL = "SELECT adminID, name, pwd FROM "&msdbprefix&"admin WHERE name = '"&strUsername&"' AND pwd = '"&strEncrPassword&"'"
    rsCommon.Open strSQL, Conn, 3, 3, &H0001
    If Not rsCommon.EOF Then 
      ' username and password match - this is a valid user
		  Response.Cookies("EZGBAdmin")("adminID") = rsCommon("adminID")
		  Response.Cookies("EZGBAdmin")("name") = DBDecode(strUsername)
		  Response.Cookies("EZGBAdmin").Expires = "Jan 18, 2038"
    
      Response.Redirect "admin.asp"

    Else
      Response.Redirect "admin_login.asp"
    End If
    rsCommon.Close: Set rsCommon = Nothing
    Conn.Close: Set Conn = Nothing

  End If
%>
<!DOCTYPE HTML>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>EZGuestbook - Admin Login</title>
<link type="text/css" rel="stylesheet" href="../assets/css/main.css" />
</head>

<body>
<br /><br /><br /><br />
  <div id="main" class="container" align="center" style="margin-top:-75px;">
    <div class="row 50%">
      <div class="12u 12u$(medium)">
        <header><h2>EZGuestbook Admin Login</h2></header>
      </div>
    </div>
  </div>
  <div id="main" class="container" align="center" style="margin-top:-75px;">
    <div class="row 50%">
      <div class="12u 12u$(medium)">
  
        <form action="admin_login.asp" method="POST">
        <div class="row">
          <div class="-4u 4u 12u$(medium)" style="padding-bottom:20px;">
            <label for="name">Name</label>
            <input type="text" id="name" name="name" required>
          </div>
          <div class="4u 1u$"><span></span></div>

          <div class="-4u 4u 12u$(medium)" style="padding-bottom:20px;">
            <label for="pwd">Password</label>
            <input type="password" id="pwd" name="pwd" required>
          </div>
          <div class="4u 1u$"><span></span></div>

          <div class="12u 12u$(medium)">
            <input class="button" type="submit" value="Let me in!">
          </div>
        </form>
      </div>
    </div>
  </div>
</body>
</html>