<% Option Explicit 
%>
<!--#Include File="config.asp"-->
<!--#Include File="conn.asp"-->
<!--#Include File="core.asp"-->
<!--#Include File="clsConfig.asp"-->
<!--#Include File="clsUpload.asp"-->
<%
Server.ScriptTimeout = 9000

Dim intSetting : intSetting = Request.Querystring("s")

Dim strAboutClass

If Request.Cookies("about") = "HIDE" Then
	strAboutClass = "Display: none;"
End If

Dim objConfig : Set objConfig = New MTConfig
With objConfig
	.Database		= aryMTDB
	.Setting		= intSetting
	.Config			= aryMTConfig
End With

Call CreateDatabaseConnection(1)
Dim blnAdmin : blnAdmin = Authenticate(True, aryMTDB(5))
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Stats</title>
	<link rel="stylesheet" href="style.css" type="text/css">
	<script language="JavaScript" src="javascript.js" type="text/javascript"></script>
</head>

<body>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
  <tr class=pgheader>
  	<td colspan=3>
	  <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
    	<td rowspan="2" width="267"><img src="images/FJstats_logo.gif" width="267" height="44"></td>
   		 <td colspan="6" ><img src="images/blue_scale.gif" width="2" height="23"></td>
   		</tr>
		<tr>
	    <td width="55"><img src="images/nav_pointer.gif" width="55" height="21"></td>
	    <td width="81"><a href="default.asp" onMouseOver="MM_swapImage('Image1','','images/reports_ovr.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="images/reports.gif" name="Image1" width="81" height="21" border="0" id="Image1"></a></td>
	    <% If blnAdmin = True Then %><td width="88"><a href="settings.asp" onMouseOver="MM_swapImage('Image2','','images/settings_ovr.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="images/settings.gif" name="Image2" width="88" height="21" border="0" id="Image2"></a></td><% End If %>
	    <td width="86"><a href="tracking.asp" onMouseOver="MM_swapImage('Image3','','images/tracking_ovr.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="images/tracking.gif" name="Image3" width="86" height="21" border="0" id="Image3"></a></td>
	    <td width="75"><a href="login.asp?action=logout" onMouseOver="MM_swapImage('Image4','','images/logout_ovr.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="images/logout.gif" name="Image4" width="75" height="21" border="0" id="Image4"></a></td>
	    <td background="images/nav_scale.gif"><img src="images/nav_scale.gif" width="2" height="21"></td>
	  </tr>
      </table>
	</td>
  </tr>
  <tr>
    <td colspan=3>
	  <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="61"><img src="images/subnav_pointer.gif" width="61" height="22"></td>
        <td width=258 background="images/subnav_scale.gif" valign=middle>
		<span class=sitename><% Response.Write objConfig.SiteName %></span></td>
        <td background="images/subnav_scale.gif" valign=middle align=right>
		<span class=version><% Response.Write objConfig.Version %></span></td>
		</td>
      </tr>
    </table>
  </td>
</tr>
<tr valign=top height="100%">
	<td style="padding: 5px;" width=180>
		<table border=0 cellpadding=0 cellspacing=0 class=select width=180>
		<tr>
			<td>
				<table border=0 cellpadding=0 cellspacing=0 width="100%">
				<tr>
					<td>
						<table border=0 cellpadding=0 cellspacing=0 width="100%">
						<tr>
							<td width="20"><img src="images/grey_arrow.gif"></td>
							<td class=header>Options</td>
						</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td class=chooser>
						<table cellpadding=3 cellspacing=0 border=0>
						<tr>
							<td width=20><img src="images/lt_grey_arrow_dn.gif" border=0></td>
							<td><a href="settings.asp?s=0" class=chtitle>General</a></td>
						</tr>
						<tr>
							<td width=20><img src="images/lt_grey_arrow_dn.gif" border=0></td>
							<td><a href="settings.asp?s=1" class=chtitle>Configuration</a></td>
						</tr>
						<tr>
							<td width=20><img src="images/lt_grey_arrow_dn.gif" border=0></td>
							<td><a href="settings.asp?s=2" class=chtitle>Users</a></td>
						</tr>
						<tr>
							<td width=20><img src="images/lt_grey_arrow_dn.gif" border=0></td>
							<td><a href="settings.asp?s=3" class=chtitle>Maintenance</a></td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		</table>
	</td>
	<td align=left style="padding: 5px;" width="100%">
	<%
	' GENERATE REPORT AND CALCULATE EXECUTION TIME
	Call objConfig.GenerateSettings(intSetting)
	%>
	<br><img src="images/spacer.gif" width=250 height=1>
	</td>
	<td width=215 class=about id=about style="<%=strAboutClass%>"><% Response.Write ShowProductInfo() %>
	</td>
</tr>
<tr class=pgfooter>
	<td></td>
	<td colspan=2 valign=middle align=right>
	</td>
</tr>
<tr class=pgbottom>
	<td colspan=3 height=4>&nbsp;</td>
</tr>
</table>
</body>
</html>
<%
Set objConfig = Nothing
Call CloseDatabaseConnection() 
%>