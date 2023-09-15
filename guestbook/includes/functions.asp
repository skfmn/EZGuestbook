<%
	on error resume next
	strVersion = "4"

    ConnStr = "Provider=sqloledb;Data Source="&msdbserver&";Initial Catalog="&msdb&";User Id="&msdbid&";Password="&msdbpwd

	Set Conn = Server.CreateObject("ADODB.Connection")
	Call ConnOpen(Conn)

	Set rsCommon = Server.CreateObject("ADODB.Recordset")

	Call getTableRecordset(msdbprefix&"settings",rsCommon)
	If Not rsCommon.EOF Then
		strSitetitle = rsCommon("site_title")
		strDomain = rsCommon("domain_name")
	End If
	Call closeRecordset(rsCommon)
	Call ConnClose(Conn)

	Function getResponse(sURL)
		Dim strTemp
		strTemp = ""

		Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		xmlhttp.SetOption(2) = (xmlhttp.GetOption(2) - SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS)
		xmlhttp.Open "GET", sURL, false
		xmlhttp.Send(NULL)
		If xmlhttp.readyState = 4 Then
			If xmlhttp.status <> 200 Then

			    strTemp = "<span style=""color:#FF0000"">Error "&xmlhttp.status&" - "&xmlhttp.statusText&"</span><br />"

			Else

			    strTemp = xmlhttp.ResponseText

			End If
		End If
		Set xmlhttp = Nothing

		getResponse = strTemp

	End Function

	Sub pagination(iPage)

        Response.Write "            <div class=""-3u 8u$ 12u(medium)"" style=""margin-bottom:0;"">"&vbcrlf
        Response.Write "                <ul class=""actions"" style=""margin-bottom:0;"">"&vbcrlf
        Response.Write "                    <li>"&vbcrlf

        If PageNo > 1 Then
            Response.Write "                        <ul class=""actions"">"&vbcrlf
            Response.Write "                            <li><a style=""color:#ffffff;"" class=""button fit"" href=""view.asp?page=1""><i class=""fa fa-angle-double-left""></i></a></li>"&vbcrlf
            Response.Write "                            <li><a style=""color:#ffffff;"" class=""button fit"" href=""view.asp?page="&PageNo - 1&"""><i class=""fa fa-angle-left""></i></a></li>"&vbcrlf
            Response.Write "                        </ul>"&vbcrlf
        End If

        Response.Write "                    </li>"&vbcrlf
        Response.Write "                    <li>"&vbcrlf
        Response.Write "                        <ul class=""actions"">"&vbcrlf
	 
        If  Cint(PageNo) = 1 Then
            Response.Write "                            <li><span class=""button fit""><span style=""color:lightblue;"">1</span></span</li>"&vbcrlf
        End If

        If  Cint(PageNo) = 2 Then
            Response.Write "                            <li><a style=""color:#ffffff;"" class=""button fit"" href=""view.asp?page=1"">1</a></li>"&vbcrlf
        End If

        If  Cint(PageNo) > 2 Then
            Response.Write "                            <li><a style=""color:#ffffff;"" class=""button fit"" href=""view.asp?page="&PageNo - 2&""">"&PageNo - 2&"</a></li>"&vbcrlf
            Response.Write "                            <li><a style=""color:#ffffff;"" class=""button fit"" href=""view.asp?page="&PageNo - 1&""">"&PageNo - 1&"</a></li>"&vbcrlf
        End If

        If  Cint(PageNo) <> 1 And  Cint(PageNo) <> Cint(TotalPages) Then
            Response.Write "                            <li><span class=""button fit""><span style=""color:lightblue;"">"&PageNo&"</span></span></li>"&vbcrlf
        End If

        If Cint(PageNo) < CInt(TotalPages-1) And Cint(PageNo) <> Cint(TotalPages) Then
            Response.Write "                            <li><a style=""color:#ffffff;"" class=""button fit"" href=""view.asp?page="&PageNo + 1&""">"&PageNo + 1&"</a></li>"&vbcrlf
            Response.Write "                            <li><a style=""color:#ffffff;"" class=""button fit"" href=""view.asp?page="&PageNo + 2&""">"&PageNo + 2&"</a></li>"&vbcrlf
        End If

        If  Cint(PageNo) = Cint(TotalPages)-1 Then
            Response.Write "                             <li><a style=""color:#ffffff;"" class=""button fit"" href=""view.asp?page="&TotalPages&""">"&TotalPages&"</a></li>"&vbcrlf
        End If

        If  Cint(PageNo) = Cint(TotalPages) AND  Cint(PageNo) <> 1 Then
            Response.Write "                             <li><span class=""button fit""><span style=""color:lightblue;"">"&TotalPages&"</span></span></li>"&vbcrlf
        End If

        Response.Write "                        </ul>"&vbcrlf
        Response.Write "                    </li>"	&vbcrlf
        Response.Write "                    <li>"&vbcrlf

        If Cint(PageNo) >= 1 AND Cint(TotalPages) > 1 AND Cint(PageNo) <> Cint(TotalPages) Then
            Response.Write "                        <ul class=""actions"">"&vbcrlf
            Response.Write "                            <li><a style=""color:#ffffff;"" class=""button fit"" href=""view.asp?page="&PageNo+1&"""><i class=""fa fa-angle-right""></i></a></li>"&vbcrlf
            Response.Write "                            <li><a style=""color:#ffffff;"" class=""button fit"" href=""view.asp?page="&TotalPages&"""><i class=""fa fa-angle-double-right""></i></a></li>"&vbcrlf
            Response.Write "                        </ul>"&vbcrlf
        End If

        Response.Write "                    </li>"&vbcrlf
        Response.Write "                </ul>"&vbcrlf
        Response.Write "            </div>"&vbcrlf

	End Sub

	Sub adminPagination(iPage)

        Response.Write "            <div class=""-3u 8u$ 12u(medium)"" style=""margin-bottom:0;"">"&vbcrlf
        Response.Write "                <ul class=""actions"" style=""margin-bottom:0;"">"&vbcrlf
        Response.Write "                    <li>"&vbcrlf

        If PageNo > 1 Then
            Response.Write "                        <ul class=""actions"">"&vbcrlf
            Response.Write "                            <li><a style=""color:#ffffff;"" class=""button fit"" href=""admin_entries.asp?page=1""><i class=""fa fa-angle-double-left""></i></a></li>"&vbcrlf
            Response.Write "                            <li><a style=""color:#ffffff;"" class=""button fit"" href=""admin_entries.asp?page="&PageNo - 1&"""><i class=""fa fa-angle-left""></i></a></li>"&vbcrlf
            Response.Write "                        </ul>"&vbcrlf
        End If

        Response.Write "                    </li>"&vbcrlf
        Response.Write "                    <li>"&vbcrlf
        Response.Write "                        <ul class=""actions"">"&vbcrlf
	 
        If  Cint(PageNo) = 1 Then
            Response.Write "                            <li><span class=""button fit""><span style=""color:lightblue;"">1</span></span</li>"&vbcrlf
        End If

        If  Cint(PageNo) = 2 Then
            Response.Write "                            <li><a style=""color:#ffffff;"" class=""button fit"" href=""admin_entries.asp?page=1"">1</a></li>"&vbcrlf
        End If

        If  Cint(PageNo) > 2 Then
            Response.Write "                            <li><a style=""color:#ffffff;"" class=""button fit"" href=""admin_entries.asp?page="&PageNo - 2&""">"&PageNo - 2&"</a></li>"&vbcrlf
            Response.Write "                            <li><a style=""color:#ffffff;"" class=""button fit"" href=""admin_entries.asp?page="&PageNo - 1&""">"&PageNo - 1&"</a></li>"&vbcrlf
        End If

        If  Cint(PageNo) <> 1 And  Cint(PageNo) <> Cint(TotalPages) Then
            Response.Write "                            <li><span class=""button fit""><span style=""color:lightblue;"">"&PageNo&"</span></span></li>"&vbcrlf
        End If

        If Cint(PageNo) < CInt(TotalPages-1) And Cint(PageNo) <> Cint(TotalPages) Then
            Response.Write "                            <li><a style=""color:#ffffff;"" class=""button fit"" href=""admin_entries.asp?page="&PageNo + 1&""">"&PageNo + 1&"</a></li>"&vbcrlf
            Response.Write "                            <li><a style=""color:#ffffff;"" class=""button fit"" href=""admin_entries.asp?page="&PageNo + 2&""">"&PageNo + 2&"</a></li>"&vbcrlf
        End If

        If  Cint(PageNo) = Cint(TotalPages)-1 Then
            Response.Write "                             <li><a style=""color:#ffffff;"" class=""button fit"" href=""admin_entries.asp?page="&TotalPages&""">"&TotalPages&"</a></li>"&vbcrlf
        End If

        If  Cint(PageNo) = Cint(TotalPages) AND  Cint(PageNo) <> 1 Then
            Response.Write "                             <li><span class=""button fit""><span style=""color:lightblue;"">"&TotalPages&"</span></span></li>"&vbcrlf
        End If

        Response.Write "                        </ul>"&vbcrlf
        Response.Write "                    </li>"	&vbcrlf
        Response.Write "                    <li>"&vbcrlf

        If Cint(PageNo) >= 1 AND Cint(TotalPages) > 1 AND Cint(PageNo) <> Cint(TotalPages) Then
            Response.Write "                        <ul class=""actions"">"&vbcrlf
            Response.Write "                            <li><a style=""color:#ffffff;"" class=""button fit"" href=""admin_entries.asp?page="&PageNo+1&"""><i class=""fa fa-angle-right""></i></a></li>"&vbcrlf
            Response.Write "                            <li><a style=""color:#ffffff;"" class=""button fit"" href=""admin_entries.asp?page="&TotalPages&"""><i class=""fa fa-angle-double-right""></i></a></li>"&vbcrlf
            Response.Write "                        </ul>"&vbcrlf
        End If

        Response.Write "                    </li>"&vbcrlf
        Response.Write "                </ul>"&vbcrlf
        Response.Write "            </div>"&vbcrlf

	End Sub

	Function getMessage(sMsg)
		Dim rsTemp, strTemp
		strTemp = ""
		
		Set Conn = Server.CreateObject("ADODB.Connection")
		Set rsTemp= Server.CreateObject("ADODB.Recordset")

		strSQL = "SELECT message FROM "&msdbprefix&"messages WHERE msg = '"&sMsg&"'"
		Call ConnOpen(Conn)
		Call getTextRecordset(strSQL,rsTemp)
		If Not rsTemp.EOF Then
		    strTemp = DBDecode(rsTemp("message"))
		Else
		    strTemp = sMsg
		End If
		Call closeRecordset(rsTemp)
		Call ConnClose(Conn)
		
		getMessage = strTemp
		 
	End Function

	Function msgTrans(sMsg)

		Dim strTemp: strTemp = ""

		Select  Case sMsg
			case "lic"
				strTemp = "Login info updated:"
			case "fch"
				strTemp = "Fields updated:"
			case "opa"
				strTemp = "Options updated:"
			case "dls"
				strTemp = "Site deleted:"
			case "dlf"
				strTemp = "Find deleted:"
			case "msgd"
				strTemp = "Entry deleted:"
			case "ipd"
				strTemp = "IP deleted:"
			case "ban"
				strTemp = "IP banned:"
			case "unban"
				strTemp = "IP unbanned:"
			case "sus"
				strTemp = "Thanks:"
			case "ed"
				strTemp = "Entry deleted:"
			case "aeo"
				strTemp = "Error notice:"
			case "ened"
				strTemp = "Entry updated:"
			case "mus"
			    strTemp = "Message updated:"
			case "siu"
			    strTemp = "Site updated:"
			case "cpwds"
			    strTemp = "Admin updated:"
			case "nadmin"
			    strTemp = "Not an Admin:"
			case "das"
			    strTemp = "Admin deleted:"
		End Select

		msgTrans = strTemp

	End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function DBEncode(DBvalue)
' SQL injection function for database inputs
' Reverses the input to make a hack query meaningless
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function DBDecode(DBvalue)
' SQL injection Function for database inputs
' reverses back to the original for the DB output.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
	Function DBDecode(DBvalue)
		Dim fieldvalue
		fieldvalue = Trim(DBvalue)
		
		If fieldvalue <> "" AND ( NOT IsNull(fieldvalue) ) Then
		
			Set encodeRegExp = New RegExp 
			encodeRegExp.Pattern = "((eteled)*(tceles)*(etadpu)*(otni)*(pord)*(tresni)*(eralced)*(_px)*(noinu)*)"
			encodeRegExp.IgnoreCase = True
			encodeRegExp.Global = True
			Set Matches = encodeRegExp.Execute(fieldvalue)
			For Each Match In Matches
			    fieldvalue = Replace(fieldvalue,Match.Value,StrReverse(Match.Value))
			Next
			fieldvalue = replace(fieldvalue,"''","'")

		End If
		
		DBDecode = fieldvalue
		
	End Function
	
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 

  Function RemoveHTML(strText)
	  Set RegEx = New RegExp
	  RegEx.Pattern = "<[^>]*>"
	  RegEx.Global = True
	  strText = Replace(LCase(strText), "<br>", chr(10))
		strText = RegEx.Replace(strText, "")
	  RemoveHTML = strText 
  End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function displayFancyMsg(sText)
' displays an error message for parts of this script using Fancybox
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
	Function displayFancyMsg(sText)
%>
    <div style="display:none">
		<a id="textmsg" href="#displaymsg">Message</a>
		<div id="displaymsg" style="background-color:#ffffff;text-align:left;width:300px;">
			<div class="left_menu_block">
				<div class="left_menu_top"><h2>Message</h2></div>
				<div class="left_menu_center" style="text-align:center;background-color:#ffffff; padding-left:0px;">
				    <span style="color:#000000;"><%= sText %></span>
				</div>
				<div class="left_menu_bottom"></div>
			</div>
		</div>
    </div>
<%
	End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function checkInt(sText)
' checks that it is indeed a number
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
	Function checkInt(iVal)
		Dim intTemp
		intTemp = 0
		
		If iVal <> "" Then
			If Not IsNumeric(iVal) Then
			    Call displayFancyMsg("Input was not a number!")
			Else
			    intTemp = iVal		
			End If
		Else
		    Call displayFancyMsg(txtInputWasEmpty&"Input was empty!")
		End If
		
		checkInt = intTemp
		
	End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Function fileExt(s)  
	    fileExt = Right(s,Len(s)-(inStrRev(s,"\",-1,1))) 
	End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
	Function ConvBytes(TBytes)
		Dim inSize, isType
		Const lnBYTE = 1
		Const lnKILO = 1024                     ' 2^10
		Const lnMEGA = 1048576                  ' 2^20
		Const lnGIGA = 1073741824               ' 2^30
		Const lnTERA = 1099511627776
	
		If TBytes < 0 Then Exit Function   
		If TBytes < lnKILO Then
		    ' ByteConversion
		    inSize = TBytes
		    isType = "bytes"
		Else        
			If TBytes < lnMEGA Then ' KiloByte Conversion
			    inSize = (TBytes / lnKILO)
			    isType = "kb"
			ElseIf TBytes < lnGIGA Then ' MegaByte Conversion
			    inSize = (TBytes / lnMEGA)
			    isType = "mb"
			ElseIf TBytes < lnTERA Then ' GigaByte Conversion
			    inSize = (TBytes / lnGIGA)
			    isType = "gb"
			Else ' TeraByte Conversion
			    inSize = (TBytes / lnTERA)
			    isType = "tb"
			End If
		End If

		' format number to 2 decimal places
		inSize = FormatNumber(inSize,2)
		' Return the results
		ConvBytes = inSize & " " & isType

	End Function
	
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function chkEmail(sEmail)
' Check if its a valid email address format
' **does not check if the email address exists**
' Returns: true|false
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Function chkEmail(sEmail)
		Set objRegExp = New RegExp
		searchStr = "^[\w-\.]{1,}\@([\da-zA-Z-]{1,}\.){1,}[\da-zA-Z-]{2,3}$" 
		objRegExp.Pattern = searchStr
		objRegExp.IgnoreCase = true
		chkEmail = objRegExp.Test(sEmail)
	End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Clean User Name - Clean it of unwanted
' Characters and make sure its datbase safe
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Function cleanInput(sInput)
		strNewInput = sInput

		If Instr(strNewInput,"''") > 0 Then
		    strNewInput = Replace(strNewInput,"''","")
		End If
	
		strNewInput = cleanString(strNewInput)
	
		strNewInput = GetDBfield(strNewInput)
	
		cleanInput = strNewInput
	End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Clean string of unwanted characters
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Function cleanString(strInput)

		Set regEx = New RegExp

		regEx.Pattern = "[`¬¦!£$%^&*()+=#~@;:?/><,/*|\}{]"
		regEx.IgnoreCase = True
		regEx.Global = True

		cleanString = regEx.Replace(strInput, "")
	End Function
''''''''''''''''''''''''''''''''''''''''''''''

	Sub selectRate(iRate)

		Response.Write "<select id=""rate"" name=""rate"" style=""color:#bbbbbb;"">"&vbcrlf
		Response.Write "  <option value="""">Rate Our Site</option>"&vbcrlf
		For x = 10 to 1 STEP -1
			If Cint(iRate) = Cint(x) Then
			    Response.Write "  <option value="""&x&""" selected>"&x&"</option>"&vbcrlf
			Else
			    Response.Write "  <option value="""&x&""">"&x&"</option>"&vbcrlf
			End If
		Next
		Response.Write "</select>"&vbcrlf
	End Sub

  Sub selectSite(sSite)
    Dim rsTemp
	  
	  Set rsTemp = Server.CreateObject("ADODB.Recordset")
		Call getTablerecordset(msdbprefix&"site",rsTemp)
	  If Not rsTemp.EOF Then
	    Response.Write "<select id=""site"" name=""site"" style=""color:#bbbbbb;"">" & vbcrlf
      Response.Write "<option value="""">Site Visited</option>" & vbcrlf
	    Do While Not rsTemp.EOF

        strSiteName = ""
        strSiteName = rsTemp("site_name")
        If sSite = strSiteName Then
          Response.Write "<option value="""&strSiteName&""" selected>"&strSiteName&"</option>" & vbcrlf
        Else
	        Response.Write "<option value="""&strSiteName&""">"&strSiteName&"</option>" & vbcrlf
        End If

				rsTemp.MoveNext
				If rsTemp.EOF Then Exit Do
		  Loop
	    Response.Write "</select>" & vbcrlf
	  Else
	    Response.Write "<select><option>No Options</option></select>"
	  End If
	  Call closeRecordset(rsTemp)
		
  End Sub
  
  Sub selectFind(sFind)
    Dim rsTemp
	
	  Set rsTemp = Server.CreateObject("ADODB.Recordset")
		Call getTableRecordset(msdbprefix&"find",rsTemp)
	  If Not rsTemp.EOF Then
	    Response.Write "<select id=""find"" name=""find"" style=""color:#bbbbbb;"">"&vbcrlf
      Response.Write "  <option value="""">How did you find us?</option>"&vbcrlf
	    Do While Not rsTemp.EOF
        strFind = rsTemp("find_name")
        If sFind = strFind Then
          Response.Write "<option value="""&strFind&""" selected>"&strFind&"</option>"&vbcrlf
        Else
	        Response.Write "<option value="""&strFind&""">"&strFind&"</option>"&vbcrlf
        End If
		    rsTemp.MoveNext
				If rsTemp.EOF Then Exit Do
		  Loop
	    Response.Write "</select>" & vbcrlf
	  Else
	    Response.Write "<select><option>No Options</option></select>"
	  End If
	  Call closeRecordset(rsTemp)

  End Sub

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
	
%>	