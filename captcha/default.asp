<!--#INCLUDE FILE="_captcha.asp"-->
<html>

 <head>
  <title>Captcha ASP script - prevents automated form sumbissions</title>
	<style> 
    body { font-family: arial,sans-serif; } 
    h1   { font-size: 20; } 
    h2   { font-size: 16; } 
   .item a { margin-left:3px; padding-left: 0px; padding-right: 0px; border:0px solid red; text-decoration:none; }
   .item { padding-left: 3px; padding-right: 3px; margin:0px; border:1px solid #DDEEFF;}
  </style>
 </head>

 <body>

<h1>Captcha ASP script - prevents automated form sumbissions</h1>

<div style="width:600; border:1px solid silver;padding:10px">

<div style="float:right;width:200">
<% CaptchaForm %>

<div style="border:1px solid grey; margin-left:10px; padding: 3px; margin-top:10">

 <div class="item">
  Links
 </div>
 <div class="item">
  <a class="item" href="http://www.motobit.com/util/captcha/aspcaptcha.zip" title=".zip files with this script">Download this project</a>
 </div>
 <div class="item">
  <a class="item" href="http://www.motobit.com/util/captcha/" title="Captcha ASP image test against automated form submission">Captcha for ASP home</a>
 </div>
 <div class="item">
  <a class="item" href="http://www.motobit.com/util/captcha/default.asp?progress" title="">How it works</a>
 </div>

 <div class="item">
  External Links
 </div>
 <div class="item">
  <a class="item" href="http://www.captcha.net/">Original captcha project</a>
 </div>
 <div class="item">
 
  <a class="item" href="http://www.google.com/search?q=office+web+components" title="Download office webcomponents">OWC download</a>
 </div>
 <div class="item">
  <a class="item" href="http://www.google.com/search?q=captcha">Captcha on Google</a>
 </div>


<div class="item">
  Other products
</div>
 <div class="item">
  <a class="item" href="http://pure.aspupload.net/" title="Free Upload with progress.">Pure ASP upload</a> 
 </div>
 <div class="item">
  <a class="item" href="http://www.iismonitor.net/" title="Online IIS monitor.">IIS monitor</a>
 </div>
 <div class="item">
  
 </div>

 </div>

</div>

<%
if request.querystring="progress" then
  ProgressProject
else
  HomePage
end if
%>

<div style="border-top: 1px solid silver; margin-top:10; ">
&copy; 2006 <a href="http://www.foller.cz" style="text-decoration:none;color:black">Antonin Foller</a>, <a href="http://www.motobit.com/">Motobit Software</a>, help{at}pstruh.cz
<div>


</div>

</body>
</html>




<%Sub HomePage%>
&nbsp;&nbsp;&nbsp;&nbsp; The main idea of a <a href="http://www.captcha.net/">Captcha</a> test 
(completely automated public Turing test to tell computers and humans apart) is to prevent automated form submission. 
<br />&nbsp;&nbsp;&nbsp;&nbsp; This project is a very short VB Script source code, which let's you generate
gif, jpg or png images with text which most people can easily read and enter to a form field, 
but automated computer program cannot.

<h2>About this project</h2>
&nbsp;&nbsp;&nbsp;&nbsp; Some weeks ago I had really need this simple test for ASP (and ASP.Net also), but I found
only some very long codes with css or some script generating bmp images. I have decided to create some code 
with a public object, usually installed on hostings. One of them is a Chart object from 
<a href="http://www.google.com/search?q=office+web+components">Office web components</a>. 


<br /><br />&nbsp;&nbsp;&nbsp;&nbsp; So I created a short include generating a "human testing" images and
some form samples written in ASP. 
You can <a href="http://www.motobit.com/util/captcha/aspcaptcha.zip" 
title=".zip files with this script">download them from motobit.com</a> site. 
The code is free. 
Please include a next link on your pages if you use this include or samples:
<pre>
This site is using
&lt;a href="http://www.motobit.com/util/captcha/" 
  title="Captcha ASP image test against automated form 
  submission"&gt;Captcha for ASP&lt;/a&gt;
</pre>


<h2>Using this VBScript code and includes</h2>

&nbsp;&nbsp;&nbsp;&nbsp; This project contains one ASP include, _captcha.asp 
and several samples. Please take a look at sample1 folder, form-session.asp sample.

<br />&nbsp;&nbsp;&nbsp;&nbsp; form-session.asp generates a random text (using RandomText function from _captcha.asp), 
which is stored to a session. 
<pre>
session("checktext") = RandomText(5)
</pre>

 and the source form in form-session.asp contains one more text field named imagecheck:
<pre>
&lt;input type="text" name="imagecheck" size="5" /&gt;
</pre>

Generate-captcha.asp script creates a GIF image with the random text, using textToGIF function and BinaryWrite method.

<pre>
&lt;!--#INCLUDE FILE="_captcha.asp"--&gt;
&lt;%
  response.ContentType = "image/gif"
  response.binarywrite textToGIF(session("checktext"))
%&gt;
</pre>

form-session.asp accepts form dat then and checks validity of text written by client.
<pre>
if len(request("imagecheck"))>0 and _
  ucase("" & request("imagecheck")) = ucase("" & session("checktext")) then 
    response.write "The form was accepted."
end if
</pre>

&nbsp;&nbsp;&nbsp;&nbsp; You can see more technical details on <a class="item" href="http://www.motobit.com/util/captcha/default.asp?progress" title="">How it works</a> page.

<h2>Install office web components</h2>

<div style="">
1. <a href="http://www.google.com/search?q=office+web+components">download</a> Office web components from MS site.
<br>2. Install the software on your server (run owc10.exe or owc11.exe on your server)
or install it manually. 
<br>&nbsp;&nbsp;&nbsp; - unpack the owc11.exe  file using 'owc11.exe /C /T:c:\temp
<br>&nbsp;&nbsp;&nbsp; - unpack OWC11.DLL (OWC10.DLL) from the msi package
<br>&nbsp;&nbsp;&nbsp; - copy OWC11.DLL (OWC10.DLL) to your server and register the DLL by 'regsvr32 OWC11.DLL'

<br><br>Remember that the Office web components are licensed. Next is a part of OWC license agreement:
<br><br>1.  GRANT OF LICENSE.   If You licensed the Software from Your hardware manufacturer, You may install and use one (1) copy of the Software.  If You have a valid end user license for Microsoft Office 2003 or any component application of the Microsoft Office System, Microsoft Office XP,  Access 2002, Excel 2002, FrontPage 2002, or any other product identified by Microsoft and with which the Software interoperates (the "Licensed Products"), and You licensed the Software from Microsoft, You may use one (1) copy of the Software in accordance with the end user license agreement that accompanied the Licensed Product.  If You are not a licensee of any of the Licensed Products, You may only install and use one (1) copy of the Software for the sole purpose of viewing and printing copies of static documents, text and images created with the Software; You cannot make any other use of the Software whatsoever.
</div>

<%End Sub %>






<%Sub ProgressProject%>

<h2>Progress of this project</h2>

&nbsp;&nbsp;&nbsp;&nbsp; Some weeks ago I had really need this simple test for ASP (and ASP.Net also), but I found
only some very long codes with css or some script generating bmp images. I have decided to use some simple code 
with public object, Chart from <a href="http://www.google.com/search?q=office+web+components">Office web components</a>. 
The component works on most Windows 2000 or 2003 hostings (please see licensing note on the download
page if you have your own hosting)

<h2>How it works inside</h2>

Some of such object is a Chart object from 
<a href="http://www.google.com/search?q=office+web+components">Office web components</a>. 
You can generate a gif picture in a short code using the component:

<pre>
  Dim Chs
  Set Chs = CreateObject("OWC10.ChartSpace") 

  Chs.ExportPicture "c:\temp\image.gif", , 120, 55
</pre>

&nbsp;&nbsp;&nbsp;&nbsp; And there is a quick way to captcha image - simple add a chart, 
create a title (HasTitle = True), write some text to the title (Caption = inText), set 
font family and size and the image finished.

<pre>
  'Get chart constants
  Dim chConstants: Set chConstants = Chs.Constants
  
  'Get a chart object 
  Dim Chart: Set Chart = Chs.Charts.Add

  'Enable title for the chart.
  Chart.HasTitle = True

  'Set the text and properties.  
  Chart.Title.Caption = inText
  Chart.Title.Font.Name = "Algerian" '"Algerian"
  Chart.Title.Font.Size = 20
</pre>

&nbsp;&nbsp;&nbsp;&nbsp; The last thing to do is to set some background patterns, 
font styles, sizes and colors to make harder to read the text by automated program. You can use several 
methods of <a href="http://www.google.com/search?q=Interior+object+owc" />Interior</a> object, the best 
of them for this usage is a <a href="http://www.google.com/search?q=SetPresetGradient" />SetPresetGradient</a> method.
There are several preset patterns and colors gallery in the office web components library. You can see some of them bellow:
<br /><br />
<img alt="Human image test" src="generate-captcha.asp?text=ABCD" border="0"/>
<img alt="Human image test" src="generate-captcha.asp?text=EFGH" border="0"/>
<img alt="Human image test" src="generate-captcha.asp?text=IJKL" border="0"/>
<img alt="Human image test" src="generate-captcha.asp?text=MNOP" border="0"/>

<br /><br />&nbsp;&nbsp;&nbsp;&nbsp; The chart is created. Now you can export the "chart" (empty chart with title only) by 
<a href="http://www.google.com/search?q=ExportPicture+owc" />ExportPicture</a>
method of the chartspace object. Then I used <a href="http://www.motobit.com/tips/detpg_read-write-binary-files/">ReadBinaryFile</a>
VBS function to retrieve the stored image from a disk and BinaryWrite method to show the image on client.

<h2>Performance of the script</h2>
&nbsp;&nbsp;&nbsp;&nbsp; There is no problem with performance of the 
image generating script. I tested office web components with very high load also,
and my code using OWC can handle up to 200 requests/second on AMD 3000+. (I mean you never reach the limit :-)
<%End Sub %>










<%Sub CaptchaForm%>
<form type="post">

<div style="border:1px solid grey; margin-left:10px; padding: 10px; text-align:center; margin-top:10">
<%
	'Check if the stored session text - session("checktext") is the same as the
	' text which has client entered - request("imagecheck").
  if isempty(request("imagecheck")) then
  else
    if ucase("" & request("imagecheck")) = ucase("" & session("checktext")) then 
%>
<span style="color:green; font-weight:bold">The form was accepted. Check text: '<%=request("imagecheck")%>'</span><br />
<%

    else
%>
<span style="color:red; font-weight:bold">The image check does not pass.</span><br />
<%
    end if
  end if

	'Create a random text and store the text to session.
	' Image-Check.asp will show the text in a captcha image
	session("checktext") = RandomText(5)
%>

 
<a href="http://www.motobit.com/util/captcha/" title="Captcha ASP image test against automated form submission"><img alt="Captcha ASP image test against automated form submission" src="generate-captcha.asp" border="0"/></a>

<br />
<input type="text" name="imagecheck" size="5" /> <input type="submit" value="Test it!" />

<%if isempty(getOWC) then %>
<br /><span style="color:red;font-size:x-small" />&nbsp;&nbsp;&nbsp;&nbsp;
This script requires Office web components installed on your windows web server. 
The office web components are not installed. Please download them from 
<a href="http://www.google.com/search?q=office+web+components+download">MS site</a>
</span>
<%end if%>




</div>
</form>

<%End Sub %>
