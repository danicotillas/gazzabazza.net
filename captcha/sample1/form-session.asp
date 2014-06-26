<!--#INCLUDE FILE="_captcha.asp"-->
<html>

 <head>
  <title>Captcha ASP image test</title>
	<style> body { font-family: arial,sans-serif; } </style>
 </head>

 <body>

<h1>ASP captcha image test</h1>

<%
	'Check if the stored session text - session("checktext") is the same as the
	' text which has client entered - request("imagecheck").
	if len(request("imagecheck"))>0 and ucase("" & request("imagecheck")) = ucase("" & session("checktext")) then 
%>
	The form was accepted. Username: '<%=request("UserName")%>', Password: '<%=request("Password")%>'.
<%
	else
		'Create a random text and store the text to session.
		' Image-Check.asp will show the text in a captcha image
		session("checktext") = RandomText(5)
%>
<div style="width:400; border:1px solid silver;padding:10px">


<form type="post">

<div style="border:1px solid grey;padding:10px; margin-top:10">
 <br /><div style="width: 100; float:left; "> User Name</div><input type="text" name="UserName" value="<%=request("UserName")%>" />
 <br /><div style="width: 100; float:left; "> Password</div> <input type="text" name="Password" Type="password" value="<%=request("Password")%>" />
</div>

<div style="border:1px solid grey;padding:10px; text-align:center; margin-top:10">
<%
	if len(request("imagecheck"))>0 then response.write "The image check does not pass."
%>
 Write the text from image below to this textbox
 <br /><input type="text" name="imagecheck" size="5" />
 <br /><br />
 <a href="http://www.motobit.com/util/captcha/" ><img alt="Human image test" src="generate-captcha.asp" border="0"/></a>
</div>

<div style="border:1px solid grey; padding:10px; margin-top:10; text-align:right">
 <input type="submit" value="Send Form" />
</div>

</form>


&nbsp;&nbsp;&nbsp;&nbsp; This script is a <a href="http://www.captcha.net/">Captcha</a> test 
(completely automated public Turing test to tell computers and humans apart)
- it simply prevents automated form submission. The script written as ASP/VBS html page, see 
more about this script at its technology at <a href="http://www.motobit.com/util/captcha/" >Captcha ASP script</a> page.
<br />&nbsp;&nbsp;&nbsp;&nbsp; You can write any user name and password, but you have to enter 
the text shown in the image to the textbox bellow to send form successfully.



<%if isempty(getOWC) then %>
<br /><span style="color:red" />&nbsp;&nbsp;&nbsp;&nbsp;
This script requires Office web components installed on your windows web server. 
The office web components are not installed. Please download them from 
<a href="http://www.google.com/search?q=office+web+components+download">MS site</a>
</span>
<%end if%>


<div style="border-top: 1px solid silver; margin-top:10; ">
&copy; 2006 <a href="http://www.foller.cz" style="text-decoration:none;color:black">Antonin Foller</a>, <a href="http://www.motobit.com/">Motobit Software</a>, help{at}pstruh.cz
<div>


</div>

</body>
</html>
<%
	end if
%>
