<% Option Explicit
%>
<!--#Include File="config.asp"-->
<!--#Include File="conn.asp"-->
<!--#Include File="core.asp"-->
<%
Dim strUsername : strUsername = Request.Form("username")
Dim strPassword : strPassword = Request.Form("password")
Dim blnRemember : blnRemember = Request.Form("remember")
Dim strAction : strAction = UCase(Request("action"))

Dim strChecked, strError

If blnRemember = "ON" Then
	blnRemember = True
	strChecked = " checked"
Else 
	blnRemember = False
End If

Select Case strAction

Case "LOGIN"

	If strUsername <> "" Then
	
		Response.Cookies("FJstats")("username")	= strUsername
		Response.Cookies("FJstats")("password")	= strPassword
		
		Call CreateDatabaseConnection(1)
		Dim blnAdmin : blnAdmin = CInt(Authenticate(False, aryMTDB(5)))
		Call CloseDatabaseConnection()
		
		If blnRemember = True Then
			Response.Cookies("FJstats").expires = dateadd("d", 365, date())
		End If
		
		' REDIRECT
		Response.Redirect "default.asp"
	End If

Case "LOGOUT"

	Response.Cookies("FJstats").expires = DateAdd("d", -1, Now())
	strError = "<p class=error>You have been logged out</p>"

Case "FAILURE"
	
	Dim intCode : intCode = CInt(Request.Querystring("code"))
	If intCode = 0 Then
		strError = "<p class=error>Invalid username or password.</p>"
	ElseIf intCode = -1 Then
		strError = "<p class=error>Insufficient priviledges</p>"
	ElseIf intCode = -2 Then
		strError = "<p class=error>Please log in</p>"
	End If

End Select

' RETRIEVE USERNAME / PASSWORD FROM COOKIES
strUsername = Request.Cookies("FJstats")("username")
strPassword = Request.Cookies("FJstats")("password")

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Stats</title>
	<link rel="stylesheet" href="style.css" type="text/css">
</head>

<body style="padding-top: 50px;">

<form method=post action="login.asp">
<table border=0 cellpadding=0 cellspacing=0 class=login align=center>
<tr class=pgheader>
	<td><div align="center"><font color="#FFFFFF">FJ Stats </font></div></td>
</tr>
<% If strError <> "" Then %>
<tr>
	<td align=center style="padding: 10px;"><% Response.Write(strError) %></td>
</tr>
<% End If %>
<tr valign=top>
	<td align=center style="padding: 10px;">
	<table border=0 cellpadding=2 cellspacing=0 align=center>
	<tr>
		<td align=right><p>Username: </p></td>
		<td align=left><input type=text name=username value="<% = strUsername %>" maxlength=20 size=15></td>
	</tr>
	<tr>
		<td align=right><p>Password: </p></td>
		<td align=left><input type=password name=password value="<% = strPassword %>" maxlength=20 size=15></td>
	</tr>
	<tr valign=top>
		<td align=center colspan=2>
		<p><input type=checkbox name=remember value="ON" class=checkbox <% = strChecked %>>&nbsp;Remember login information</p>
		</td>
	</tr>
	<tr>
		<td colspan=2 align=center>
		<input type=image name=login src="images/login_btn.gif" value="Login" border=0>
		<input type=hidden name=Action Value="Login">
		</td>
	</tr>
	</table>
	</td>
</tr>
<tr class=pgfooter>
	<td align=center>&nbsp;</td>
</tr>
</table>
</form>
</body>
</html>
