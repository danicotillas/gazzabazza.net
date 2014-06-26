<% Option Explicit 
%>
<!--#Include File="config.asp"-->
<!--#Include File="conn.asp"-->
<!--#Include File="core.asp"-->
<!--#Include File="clsReport.asp"-->
<%
Dim strAboutClass, strScriptDir, strAction, blnExclude

Dim intTracking : intTracking = Request.Querystring("t")

strScriptDir = Request.Servervariables("script_name")
strScriptDir = Left(strScriptDir, InStrRev(strScriptDir, "/"))

If Request.Cookies("about") = "HIDE" Then
	strAboutClass = "Display: none;"
End If

Dim objReport : Set objReport = New MTReport
With ObjReport
	.Database			= aryMTDB
	.Config				= aryMTConfig
End With

Call CreateDatabaseConnection(1)
Dim blnAdmin : blnAdmin = Authenticate(False, aryMTDB(5))

strAction = Request.Form("action")

Select Case strAction
Case "Exclude Visits"
	Response.Cookies("tosh_exclude") = 1
	Response.Cookies("tosh_exclude").Path = "/"
	Response.Cookies("tosh_exclude").Expires = DateAdd("m", 24, Date())
Case "Include Visits"
	Response.Cookies("tosh_exclude") = ""
	Response.Cookies("tosh_exclude").Path = "/"
End Select

If Request.Cookies("tosh_exclude") <> "" Then
	blnExclude = False
Else
	blnExclude = True
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Stats</title>
	<link rel="stylesheet" href="style.css" type="text/css">
	<script language="JavaScript" src="javascript.js" type="text/javascript"></script>
</head>

<body>
<table border=0 cellpadding=0 cellspacing=0 width="100%" height="100%">
  <tr id="header" class=pgheader>
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
    <td colspan=3><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width=61><img src="images/subnav_pointer.gif" width="61" height="22"></td>
		<td width=258 background="images/subnav_scale.gif" valign=middle>
		<span class=sitename><% Response.Write objReport.SiteName %></span></td>
        <td background="images/subnav_scale.gif" valign=middle align=right>
		<span class=version><% Response.Write objReport.Version %></span></td>
		</td>
      </tr>
    </table>
  </td>
</tr>
<tr valign=top height="100%">
	<td>
		<table border=0 cellpadding=0 cellspacing=0>
		<tr>
			<td id="chooser" style="padding: 5px;" width=180>
				<table border=0 cellpadding=0 cellspacing=0 class=select width=180>
				<tr>
					<td>
						<table border=0 cellpadding=0 cellspacing=0 width="100%">
						<tr>
							<td>
								<table border=0 cellpadding=0 cellspacing=0 width="100%">
								<tr>
									<td width="20"><img src="images/grey_arrow.gif"></td>
									<td class=header>Tracking</td>
								</tr>
								</table>
							</td>
						</tr>
						<tr>
							<td class=chooser>
								<table cellpadding=3 cellspacing=0 border=0>
								<tr>
									<td width=20><img src="images/lt_grey_arrow_dn.gif" border=0></td>
									<td><a href="tracking.asp?t=0" class=chtitle>Overview</a></td>
								</tr>
								<tr>
									<td width=20><img src="images/lt_grey_arrow_dn.gif" border=0></td>
									<td><a href="tracking.asp?t=1" class=chtitle>Web Pages</a></td>
								</tr>
								<tr>
									<td width=20><img src="images/lt_grey_arrow_dn.gif" border=0></td>
									<td><a href="tracking.asp?t=3" class=chtitle>Downloads</a></td>
								</tr>
								</table>
							</td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td></td>
		</tr>
		<tr>
			<td id="chooser" style="padding: 5px;" width=180>
				<table border=0 cellpadding=0 cellspacing=0 class=select width=180>
				<tr>
					<td>
					<table border=0 cellpadding=0 cellspacing=0 width="100%">
					<tr>
						<td>
							<table border=0 cellpadding=0 cellspacing=0 width="100%">
							<tr>
								<td width="20"><img src="images/grey_arrow.gif"></td>
								<td class=header>Exclude Visits</td>
							</tr>
							</table>
						</td>
					</tr>
					<tr>
						<form method=post>
						<td class=chooser style="padding:5px;">
						<% If blnExclude = True Then %>
						<p>Your visits are being tracked.</p>
						<p class=about>If you have a dynamic IP address, click the button below to set 
						a cookie that will exclude your visits from being tracked with this 
						browser.</p>
						<div align=center><input type=submit name=action value="Exclude Visits"></div>
						<% Else %>
						<p><span class=chooser>Your visits are not being tracked.</span></p>
						<p class=about>This browser is currently being excluded from tracking. Click the button 
						below to remove this.</p>
						<div align=center><input type=submit name=action value="Include Visits"></div><br>
						<% End If %>
						</td>
						</form>
					</tr>
					</table>
			</table></td>
		</tr>
		</table>
	</td>
	<td align=left style="padding: 5px;" width="100%">
		<% Select Case intTracking %>
		<% Case 1 %>
		<table border=0 cellpadding=3 cellspacing=0>
		<tr>
			<td>
				<table border=0 cellpadding=0 cellspacing=0>
				<tr>
					<td width=22><img src="images/white_arrow.gif"></td>
					<td width="100%"><span class=name>Web Pages</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table border=0 cellpadding=5 cellspacing=0 class=settings>
				<tr>
					<td><p>Javascript tracking provides a way to track any type of page including .htm, .html, .php, .aspx, .cfm, etc.</p>					</td>
				</tr>
				<tr>
					<th align=left>Code</th>
				</tr>
				<tr>
					<td>
						<p>Use the following code to track any web page:</p>
						<div align=center>
						  <textarea name="textarea" cols=80 rows=16 readonly nowrap>
&lt;script language=&quot;JavaScript&quot;&gt;

var pagetitle = document.title; //INSERT CUSTOM PAGE NAME IN QUOTES
var scriptlocation = &quot;<% = strScriptDir %>track.asp&quot;;

var pagedata = 'mtpt=' + escape(pagetitle) + '&mtr=' + escape(document.referrer) + '&mtt=2&mts=' + window.screen.width + 'x' + window.screen.height + '&mti=1&mtz=' + Math.random();
document.write ('&lt;img height=1 width=1 ');
document.write ('src=&quot;' + scriptlocation + '?' + pagedata + '&quot;&gt;');
&lt;/script&gt;
&lt;noscript&gt;
&lt;img src=&quot;<% = strScriptDir %>track.asp?mtt=2&mti=1&quot; alt=&quot;web analytics&quot; border=0&gt;
&lt;/noscript&gt;</textarea>
						</div>
						<p>The ideal spot to place your tracking code is at the bottom of each web page before the closing html 
					tag (&lt;/html&gt;). For the page title to be automatically inserted, you must place the tracking code 
					after the &lt;/title&gt; tag and there must be text in the title tag.</p>					</td>
				</tr>
				</table>
			</td>
		</tr>
		</table>
		
		<% Case 2 %>
		
		<table border=0 cellpadding=3 cellspacing=0>
		<tr>
			<td>
				<table border=0 cellpadding=0 cellspacing=0>
				<tr>
					<td width=22><img src="images/white_arrow.gif"></td>
					<td width="100%"><span class=name>Active Server Pages (ASP)</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table border=0 cellpadding=5 cellspacing=0 class=settings>
				<tr>
					<td><p>Using ASP is the optimal way to track your web pages. It allows you to track robots accessing your 
					web pages which javascript will not. The only limitation with using ASP tracking is that you will not 
					collect screen area data from your visitors.</p></td>
				</tr>
				<tr>
					<td>
						<p>To track your .ASP files, add the following code to any .ASP page that you want to track:</p>
						<div align=center><textarea cols=60 rows=4 readonly>&lt;%
Response.Cookies(&quot;tosh&quot;)(&quot;pagetitle&quot;) = &quot;&quot;
Server.Execute(&quot;<% = strScriptDir %>track.asp&quot;)
%&gt;</textarea></div>
					</td>
				</tr>
				<tr>
					<td>
						<p>You can also use a standard include in your asp pages like this:</p>
						<div align=center><textarea cols=60 rows=4 readonly>&lt;%
Response.Cookies(&quot;tosh&quot;)(&quot;pagetitle&quot;) = &quot;&quot;
%&gt;
&lt;!-- #Include Virtual=&quot;<% = strScriptDir %>track.asp&quot; --&gt;</textarea></div>
					</td>
				</tr>		
				</table>
			</td>
		</tr>
		</table>
		
		<% Case 3 %>
		
		<table border=0 cellpadding=3 cellspacing=0>
		<tr>
			<td>
				<table border=0 cellpadding=0 cellspacing=0>
				<tr>
					<td width=22><img src="images/white_arrow.gif"></td>
					<td width="100%"><span class=name>Downloads</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table border=0 cellpadding=5 cellspacing=0 class=settings>
				<tr>
					<td><p>You can track files that aren't web pages by linking to them in a special way in your HTML page. This is useful for tracking downloads, media files, etc.</p></td>
				</tr>
				<tr>
					<td>
						<p>Here is an example of how to link to a file and track it:</p>
						<div align=center><textarea cols=80 rows=3 readonly>&lt;a href="<% = strScriptDir %>track.asp?mtr=/downloads/somefile.zip"&gt;Download&lt;/a&gt;</textarea></div>
					</td>
				</tr>
				<tr>
					<td><p>Simply change the /downloads/somefile.zip to the path and name of the file you would like to track. You can use full URLs as well such as http://www.yourdomain.com/downloads/yourfile.zip.</p></td>
				</tr>				
				</table>
			</td>
		</tr>
		</table>
		
		<% Case Else '0 %>
		<table border=0 cellpadding=3 cellspacing=0>
		<tr>
			<td>
				<table border=0 cellpadding=0 cellspacing=0>
				<tr>
					<td width=22><img src="images/white_arrow.gif"></td>
					<td width="100%"><span class=name>Overview</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table border=0 cellpadding=5 cellspacing=0 class=settings>
				<tr>
					<td><p class=about>To track activity on your web site, you will need to insert 
					tracking code on each web page. There are two ways to track your web content:</p> 
					<ul>
						<li><a href="?t=1">Web Pages</a> - Track activity on your web pages using Javascript inserted into each page</li>
						<li><a href="?t=3">Downloads</a> - Track downloads or multimedia files using a simple redirect</li>
					</ul>
					<p class=about>To setup web site tracking, follow these steps:</p>
						<ol>
							<li>Choose a tracking method based on the type of file you are tracking</li>
							<li>Copy the code and paste it into your web pages</li>
						</ol>
					</td>
				</tr>		
				</table>
			</td>
		</tr>
		</table>
		<% End Select %>
	</td>
	<td class=about width=215 id=about style="<%=strAboutClass%>"><% Response.Write ShowProductInfo() %></td>
</tr>
<tr class=pgfooter id=pgfooter>
	<td><img src="images/blue_scale.gif" width="2" height="23"></td>
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
Set objReport = Nothing
Call CloseDatabaseConnection() 
%>
