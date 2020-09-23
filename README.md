<div align="center">

## A\+ Email Verification


</div>

### Description

Will call a webservice that will verify an email address down to server level. This service is provided for FREE! No costs involved. Try it out!
 
### More Info
 
email address

This calls a web service that is free to use if you only check 1000 or less addresses

A code that tells how good the address is

Must have MSXML3.0 installed from MSDN on the server you are using ASP on.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[TinyQuote](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tinyquote.md)
**Level**          |Beginner
**User Rating**    |4.4 (35 globes from 8 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Internet/ Browsers/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-browsers-html__4-9.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tinyquote-a-email-verification__4-7300/archive/master.zip)





### Source Code

```
<html>
<%@LANGUAGE="VBScript"%>
<%
dim email
dim status
dim emaildata
if Request.Form.Count > 0 then
	' Requires Microsoft XML SDK 3.0 available at msdn.microsoft.com.
	' fill data
	email = Request.Form("email")
	' Call Webservice at CDYNE
	 Dim oXMLHTTP
	 ' Call the web service to get an XML document
	 Set oXMLHTTP = server.CreateObject("Msxml2.ServerXMLHTTP")
	 oXMLHTTP.Open "POST", _
	    "http://ws.cdyne.com/emailverify/ev.asmx/VerifyEmail", _
	    False
	 oXMLHTTP.setRequestHeader "Content-Type", _
	       "application/x-www-form-urlencoded"
	 oXMLHTTP.send "email=" & server.URLEncode(email)
	 Response.Write oxmlhttp.status
	 If oXMLHTTP.Status = 200 Then
		 Dim oDOM
		 Set oDOM = oXMLHTTP.responseXML
		 Dim oNL
		 Dim oCN
		 Dim oCC
		 Set oNL = oDOM.getElementsByTagName("ReturnIndicator")
		 For Each oCN In oNL
		 For Each oCC In oCN.childNodes
		  Select Case LCase(oCC.nodeName)
		   Case "responsetext"
		    emaildata = emaildata & "CodeTxt: " & occ.text & "<br>"
		   Case "responsecode"
		    emaildata = emaildata & "Code: " & occ.text & "<br>"
		  End Select
		 Next
		 Next
		 if status = "" then status = "OK"
		 Set oCC = Nothing
		 Set oCN = Nothing
		 Set oNL = Nothing
		 Set oDOM = Nothing
	 else
	 Status = "Service Unavailable. Try again later"
	 End If
	 Set oXMLHTTP = Nothing
end if
%>
<HEAD>
<BODY><form method="POST" action="">
 <p>Email Address Checker<BR>
 <input type="text" name="email" size="40" value="<%=email%>"></p><%=status %>
 <p><input type="submit" value="Check Email" name="B1"></p>
 <p><%=emaildata%></p>
</form></BODY>
</html>
```

