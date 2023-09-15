<% 
on error resume next
    Sub trace(strText)
	    Response.Write "Debug: "&strText&"<br />"&vbcrlf
    End Sub

    Sub catch(sText,sText2)

	    If Err.Number <> 0 then
		    Call trace(sText&" - "&err.description)
	    Else
		    Call trace(sText&" - no error")
	    End If
	    If sText2 <> "" Then
		    Call trace(sText&" - "&sText2)
	    End If

	    on error goto 0
    End Sub

    Const ForReading = 1
    Const TristateUseDefault = -2

	Function DBEncode(DBvalue)
		Dim fieldvalue
		fieldvalue = Trim(DBvalue)

		If fieldvalue <> "" AND Not IsNull(fieldvalue) Then
		
			Set encodeRegExp = New RegExp
			encodeRegExp.Pattern = "((delete)*(select)*(update)*(into)*(drop)*(insert)*(declare)*(xp_)*(union)*)"
			encodeRegExp.IgnoreCase = True
			encodeRegExp.Global = True
			Set Matches = encodeRegExp.Execute(fieldvalue)
			For Each Match In Matches
				fieldvalue = Replace(fieldvalue,Match.Value,StrReverse(Match.Value))
			Next
			fieldvalue=replace(fieldvalue,"'","''")

		End If

		DBEncode = fieldvalue

	End Function    
%>
<!DOCTYPE HTML>
<html>
<head>
<title>Install</title>
<link type="text/css" rel="stylesheet" href="../assets/css/main.css" />
</head>
<body>
  <div id="main" class="container" align="center" style="margin-top:-75px;">
    <div class="row 50%">
      <div class="12u 12u$(medium)">
        <header><h2>EZGuestbook Installation</h2></header>
      </div>
    </div>
  </div>
<% If Trim(Request.QueryString("step")) = "one" Then %>
  <div id="main" class="container" align="center" style="margin-top:-100px;">
    <div class="row 50%">
      <div class="12u 12u$(medium)">
        <form action="install.asp?setsql=y" method="post">
        
        <header>
          <h2>MSSQL Database</h2>
        </header>
        <div class="row">
          <div class="-4u 4u 12u$(medium)" style="padding-bottom:20px;">
            <label for="svrname" style="text-align:left;">Server Host Name or IP Address
              <input type="text" name="svrname" required>
            </label>
          </div>
          <div class="4u 1u$"><span></span></div>

          <div class="-4u 4u 12u$(medium)" style="padding-bottom:20px;">
            <label for="dbname" style="text-align:left;">Database Name
              <input type="text" name="dbname" required>
            </label>
          </div>
          <div class="4u 1u$"><span></span></div>

          <div class="-4u 4u 12u$(medium)" style="padding-bottom:20px;">
            <label for="dbid" style="text-align:left;">Database Login
              <input type="text" name="dbid" required>
            </label>
          </div>
          <div class="4u 1u$"><span></span></div>

          <div class="-4u 4u 12u$(medium)" style="padding-bottom:20px;">
            <label for="dbpwd" style="text-align:left;">Database Password
              <input type="password" name="dbpwd" required>
            </label>
          </div>
          <div class="4u 1u$"><span></span></div>

          <div class="-4u 4u 12u$(medium)" style="padding-bottom:20px;">
            <label for="dbprefix" style="text-align:left;">Table Prefix
              <input type="text" name="dbprefix" value="ezgbk_" required>
            </label>
          </div>
          <div class="4u 1u$"><span></span></div>

          <div class="12u 12u$(medium)">
           <input class="button" type="submit" name="submit" value="Continue">
          </div>
        </div>
        </form>     
      </div>
    </div>
  </div>
<% 
  ElseIf Request.QueryString("setsql") = "y" Then

%>
  <div id="main" class="container" align="center">
    <div class="row 50%">
      <div class="12u 12u$(medium)">
<%

    msdbserver = DBencode(Request.Form("svrname"))
    msdb = DBencode(Request.Form("dbname"))
    msdbid = DBencode(Request.Form("dbid"))
    msdbpwd = DBencode(Request.Form("dbpwd"))
    msdbprefix = DBencode(Request.Form("dbprefix"))

    Set Conn = Server.CreateObject("ADODB.Connection")
    Conn.Open "Provider=sqloledb;Data Source="&msdbserver&";Initial Catalog="&msdb&";User Id="&msdbid&";Password="&msdbpwd
		
	Set rsCommon = Server.CreateObject("ADODB.Recordset")

    Response.Write "Creating Database Tables<br /><br />"
	Response.Write "Creating admin table...<br />"
	Response.Flush

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0
	
    Conn.Execute "CREATE TABLE "&msdbprefix&"admin " & _
    "([adminID] [numeric](10, 0) IDENTITY (1, 1) CONSTRAINT [PK_"&msdbprefix&"admin] PRIMARY KEY," & _
    "[name] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
    "[pwd] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & _
    "[salt] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL)"
	 
    Response.Write "Populating admin table...<br />"
	Response.Flush

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0	

	Conn.Execute "INSERT INTO "&msdbprefix&"admin ([name],[pwd],[salt]) VALUES ('admin','EB36FB0C1F1A92A838AA1ECAAD4AB6E3B5257103','833D1')"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0

    Response.Write "Creating settings table...<br />"
    Response.Flush
	
    Conn.Execute "CREATE TABLE "&msdbprefix&"settings "& _
    "([settingID] [numeric](10, 0) IDENTITY (1, 1) CONSTRAINT [PK_"&msdbprefix&"settings] PRIMARY KEY," & _
    "[site_title] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
    "[domain_name] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL)"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0

	Response.Write "Creating Messages table...<br />"
	Response.Flush
				  
    Conn.Execute "CREATE TABLE "&msdbprefix&"messages " & _ 
    "([messageID] [numeric] IDENTITY (1, 1) CONSTRAINT [PK_"&msdbprefix&"messages] PRIMARY KEY," & _
    "[msg] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
    "[message] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL)"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0
  
    Response.Write "Populating Messages table...<br />"
	Response.Flush
	
	Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('lic','Your Login Info Has Been Changed!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('fch','Fields have been changed!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('opa','Options have been changed!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('dls','Site Deleted!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('dlf','Find Deleted!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('msgd','Entry Deleted!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('ipd','IP Deleted!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('ban','You have banned the IP address!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('unban','You have un-banned the IP Address!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('sus','Thank you for signing our guestbook!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('ed','Entry Deleted!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('aeo','The operation could not be completed!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('ened','Entry Edited!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('mus','Message updated!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('siu','Site updated!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('cpwds','Admin upated!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('nadmin','Not an Admin!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('das','Admin deleted!')"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0
    
    Response.Write "Creating Fields table...<br />"
    Response.Flush
				  
    Conn.Execute "CREATE TABLE "&msdbprefix&"fields " & _ 
    "([fieldID] [numeric] IDENTITY (1, 1) CONSTRAINT [PK_"&msdbprefix&"fields] PRIMARY KEY," & _
    "[field_name] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
    "[field_show] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL)"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0

    Response.Write "Populating Fields table...<br />"
    Response.Flush
	
    Conn.Execute "INSERT INTO "&msdbprefix&"fields([field_name],[field_show]) VALUES('Name','yes')"
    Conn.Execute "INSERT INTO "&msdbprefix&"fields([field_name],[field_show]) VALUES('Email','yes')"
    Conn.Execute "INSERT INTO "&msdbprefix&"fields([field_name],[field_show]) VALUES('Website','yes')"
    Conn.Execute "INSERT INTO "&msdbprefix&"fields([field_name],[field_show]) VALUES('facebook','yes')"
    Conn.Execute "INSERT INTO "&msdbprefix&"fields([field_name],[field_show]) VALUES('twitter','yes')"
    Conn.Execute "INSERT INTO "&msdbprefix&"fields([field_name],[field_show]) VALUES('Age','yes')"
    Conn.Execute "INSERT INTO "&msdbprefix&"fields([field_name],[field_show]) VALUES('Location','yes')"
    Conn.Execute "INSERT INTO "&msdbprefix&"fields([field_name],[field_show]) VALUES('Site Visited','yes')"
    Conn.Execute "INSERT INTO "&msdbprefix&"fields([field_name],[field_show]) VALUES('Site Rating','yes')"
    Conn.Execute "INSERT INTO "&msdbprefix&"fields([field_name],[field_show]) VALUES('Find Us?','yes')"
    Conn.Execute "INSERT INTO "&msdbprefix&"fields([field_name],[field_show]) VALUES('Comments','yes')"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0
		
    Response.Write "Creating find table...<br />"
    Response.Flush
				  
    Conn.Execute "CREATE TABLE "&msdbprefix&"find " & _ 
    "([findID] [numeric] IDENTITY (1, 1) CONSTRAINT [PK_"&msdbprefix&"find] PRIMARY KEY," & _
    "[find_name] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL)"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0
  
    Response.Write "Populating find table...<br />"
    Response.Flush
	
    Conn.Execute "INSERT INTO "&msdbprefix&"find([find_name]) VALUES('search engine')"
    Conn.Execute "INSERT INTO "&msdbprefix&"find([find_name]) VALUES('word of mouth')"
    Conn.Execute "INSERT INTO "&msdbprefix&"find([find_name]) VALUES('just surfed in')"
    Conn.Execute "INSERT INTO "&msdbprefix&"find([find_name]) VALUES('you-you idiot')"
    Conn.Execute "INSERT INTO "&msdbprefix&"find([find_name]) VALUES('Hotscripts.com')"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0

    Response.Write "Creating guestbook table...<br />"
    Response.Flush
				  
    Conn.Execute "CREATE TABLE "&msdbprefix&"guestbook " & _ 
    "([guestbookID] [numeric] IDENTITY (1, 1) CONSTRAINT [PK_"&msdbprefix&"guestbook] PRIMARY KEY," & _
    "[name] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
    "[email] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
    "[website] [nvarchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
    "[facebook] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
    "[twitter] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
    "[goggleplus] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
    "[age] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
    "[loc] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
    "[site] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
    "[rate] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
    "[find] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
    "[IP] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
    "[banIP] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
    "[gbdate] [smalldatetime] NULL, " & _
    "[comments] [nvarchar] (MAX) COLLATE SQL_Latin1_General_CP1_CI_AS NULL)"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0

    Response.Write "Creating IP table...<br />"
    Response.Flush
				  
    Conn.Execute "CREATE TABLE "&msdbprefix&"IP " & _ 
    "([IP] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL)"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0

    Response.Write "Creating options table...<br />"
    Response.Flush
				  
    Conn.Execute "CREATE TABLE "&msdbprefix&"options " & _ 
    "([orderby] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
    "[entries_page] [numeric] (10, 0) NULL ," & _ 
    "[com_count] [numeric] (10, 0) NULL)"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0

    Response.Write "Populating options table...<br />"
    Response.Flush
	
    Conn.Execute "INSERT INTO "&msdbprefix&"options([orderby],[entries_page],[com_count]) VALUES('desc',5,350)"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0

    Response.Write "Creating site table...<br />"
    Response.Flush
				  
    Conn.Execute "CREATE TABLE "&msdbprefix&"site " & _
    "([siteID] [numeric] IDENTITY (1, 1) CONSTRAINT [PK_"&msdbprefix&"site] PRIMARY KEY," & _
    "[site_name] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL)"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0

    Response.Write "Populating site table...<br />"
    Response.Flush
	
    Conn.Execute "INSERT INTO "&msdbprefix&"site([site_name]) VALUES('HTML Junction')"
    Conn.Execute "INSERT INTO "&msdbprefix&"site([site_name]) VALUES('ASP Junction')"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0
		
    Response.Write "Creating database tables...Complete!<br />"
    Response.Flush
						
    Response.Write "<br /><br />"
%>
      </div>
    </div>
  </div>
  <div id="main" class="container" align="center">
    <div class="row 50%">
      <div class="12u 12u$(medium)">
        <form action="install.asp?step=two" method="post">
        <input type="hidden" name="msdbserver" value="<%= msdbserver %>">
        <input type="hidden" name="msdb" value="<%= msdb %>">
        <input type="hidden" name="msdbid" value="<%= msdbid %>">
        <input type="hidden" name="msdbpwd" value="<%= msdbpwd %>">
        <input type="hidden" name="msdbprefix" value="<%= msdbprefix %>">
        <header>
          <h3><span class="first">You have successfully installed the MSSQL Database<br />Please click the button below to continue</span></h3>
        </header>
        <div class="row">
          <div class="12u 12u$(medium)">
           <input class="button" type="submit" name="submit" value="Continue">
          </div>
        </div>
        </form>     
      </div>
    </div>
  </div>
<%  
		Conn.Close: Set Conn = Nothing

  ElseIf Request.QueryString("step") = "two" Then  
%>
  <div id="main" class="container" align="center">
    <div class="row 50%">
      <div class="12u 12u$(medium)">
        <form action="install.asp?step=three" method="post">
        <input type="hidden" name="msdbserver" value="<%= Trim(Request.Form("msdbserver")) %>">
        <input type="hidden" name="msdb" value="<%= Trim(Request.Form("msdb")) %>">
        <input type="hidden" name="msdbid" value="<%= Trim(Request.Form("msdbid")) %>">
        <input type="hidden" name="msdbpwd" value="<%= Trim(Request.Form("msdbpwd")) %>">
        <input type="hidden" name="msdbprefix" value="<%= Trim(Request.Form("msdbprefix")) %>">
        <input type="hidden" name="PhyPath" value="<%= strPhysPath %>" />
        <header>
          <h2>Path Settings</h2>
        </header>
        <div class="row">
          <div class="-4u 4u 12u$(medium)" style="padding-bottom:20px;">
            <label for="dbid" style="text-align:left;">Base Directory
              <input type="text" name="bdir" value="<%= Request.ServerVariables("APPL_PHYSICAL_PATH") %>" />
            </label>
          </div>
          <div class="4u 1u$"><span></span></div>

          <div class="-4u 4u 12u$(medium)" style="padding-bottom:20px;">
            <label for="dir" style="text-align:left;">EZGuestbook Directory
              <input type="text" name="dir" value="/guestbook/" size="40" />
            </label>
          </div>
          <div class="4u 1u$"><span></span></div>
          <div class="12u 12u$(medium)">
           <input class="button" type="submit" name="submit" value="Continue">
          </div>
        </div>
        </form>
      </div>
    </div>
  </div>
<%
  ElseIf Request.QueryString("step") = "three" Then 

    strPageFileName = Server.MapPath("../includes/config.asp")

    Set objPageFileFSO = CreateObject("Scripting.FileSystemObject")

    If objPageFileFSO.FileExists(strPageFileName) Then
      Set objPageFileTs = objPageFileFSO.OpenTextFile(strPageFileName, 2)
    Else
      Set objPageFileTs = objPageFileFSO.CreateTextFile(strPageFileName)
    End If

    strPageEntry = Chr(60) & Chr(37) & vbcrlf & _
    "baseDir=""" & Trim(Request.Form("bdir")) & """" & vbcrlf & _
    "strDir=""" & Trim(Request.Form("dir")) & """" & vbcrlf & _
    "msdbprefix=""" & Trim(Request.Form("msdbprefix")) & """" & vbcrlf & _
    "msdbserver=""" & Trim(Request.Form("msdbserver")) & """" & vbcrlf & _
    "msdb=""" & Trim(Request.Form("msdb")) & """" & vbcrlf & _
    "msdbid=""" & Trim(Request.Form("msdbid") )& """" & vbcrlf & _
    "msdbpwd=""" & Trim(Request.Form("msdbpwd")) & """" & vbcrlf & _
    Chr(37) & Chr(62)
				 
    objPageFileTs.WriteLine strPageEntry
  
    objPageFileTs.Close

    Response.Redirect "install.asp?step=four"

  ElseIf Request.QueryString("step") = "four" Then 
%>
  <div id="main" class="container" style="margin-top:-100px;">
    <div class="row">
      <div class="12u 12u$(medium)" style="text-align:center;">
        <form action="install.asp?step=five" method="post">
        <header>
          <h2>Other stuff</h2>
        </header>
        <div class="row">

          <div class="-4u 4u 12u$(medium)" style="padding-bottom:20px;">
            <label for="sitetitle" style="text-align:left;">Site title
              <input type="text" name="sitetitle" />
            </label>
          </div>
          <div class="4u 1u$"><span></span></div>

          <div class="-4u 4u 12u$(medium)" style="padding-bottom:20px;">
            <label for="domainname" style="text-align:left;">Domain name
              <input type="text" name="domainname" value="<%= Request.ServerVariables("SERVER_NAME") %>" />
            </label>
          </div>
          <div class="4u 1u$"><span></span></div>

          <div class="12u 12u$(medium)">
           <input class="button" type="submit" name="submit" value="Continue">
          </div>
        </div>
        </form>      
      </div>
    </div>
  </div>
<%
  ElseIf Request("step") = "five" Then
    %><!-- #include file="../includes/config.asp"--><%
    Set Conn = Server.CreateObject("ADODB.Connection")
    Conn.Open "Provider=sqloledb;Data Source="&msdbserver&";Initial Catalog="&msdb&";User Id="&msdbid&";Password="&msdbpwd

    Conn.Execute "INSERT INTO "&msdbprefix&"settings ([site_title],[domain_name]) VALUES ('"&DBEncode(Request.Form("sitetitle"))&"','"&DBEncode(Request.Form("domainname"))&"')"

    Conn.Close: Set Conn = Nothing

    Response.Redirect "install.asp?step=done"

  ElseIf Request("step") = "done" Then
%>
  <div id="main" class="container">
    <div class="row">
      <div class="12u 12u$(medium)" style="text-align:center;">
        <span class="first">
          Success!
          <br>
          You have successfully configured EZGuestbook!
          <br>
          The next step is to change your password.
          <br>
          Click on the link below and login to admin.
          <br>
          Click on "Password" in the left options menu and change your password.
          <br><br>
          <a class="first" href="../admin/admin_login.asp">Login</a>
        </span>
      </div>
    </div>
  </div>
<% Else %>
  <div id="main" class="container" style="margin-top:-75px;">
    <div class="row">
      <div class="12u 12u$(medium)" style="text-align:center;">
        <span class="first">
	      You are about to install EZGuestbook.
	      <br>
	      Please follow the instructions carefully!
	      <br><br>
	      <input class="button" type="button" onClick="parent.location='install.asp?step=one'" value="Continue">
	      <br><br>
	      </span>      
      </div>
    </div>
  </div>
<% End If %>
<br />
</body>
</html>