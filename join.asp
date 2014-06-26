<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!-- #include file = "upload/config.asp" -->
	<!--#INCLUDE FILE="captcha/_captcha.asp"-->
	<!-- #include file = "Connections/gazzabazza.asp" -->
	
	





<%
if request.Form("join") <> "" then
	vJoin = request.Form("join")
	vJoin = replace(vJoin,"'","'")
	vJoin = replace(vJoin,"&#65533;","иж")
	vJoin = replace(vJoin,"&#65533;","ив")
	vJoin = replace(vJoin,"&#65533;","ик")
	vJoin = replace(vJoin,"&#65533;","ио")
	vJoin = replace(vJoin,"&#65533;","и▓")
	vJoin = replace(vJoin,"&#65533;","&ntilde;")
	vJoin = replace(vJoin,"&#65533;","ид")
	vJoin = replace(vJoin,"&#65533;","&iquest;")
	vJoin = replace(vJoin,chr(10),"<br />")
	
else
	vJoin = ""
end if
%>
<%
if vJoin <> "" AND request.Form("Submit") <> "" AND len(request("imagecheck"))>0 AND ucase("" & request("imagecheck")) = ucase("" & session("checktext")) then

		Set gazzabazza = Server.CreateObject("ADODB.Connection")
		gazzabazza.open MM_gazzabazza_STRING
		

SQL="INSERT INTO texto (ideas) values ('"& vJoin &"')"
		
		Set recordSet=gazzabazza.Execute(SQL)
	
		gazzabazza.close
		Set gazzabazza = nothing
end if

%>


<%
if vJoin <> "" AND request.Form("Submit") <> "" AND len(request("imagecheck"))>0 and ucase("" & request("imagecheck")) = ucase("" & session("checktext")) then

Dim objCDOMail
Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
objCDOMail.From = "info@gazzabazza.net"
objCDOMail.To= "danicotillas@yahoo.es"
objCDOMail.Subject= "TEXTO en Gazza&Bazza"

vMessage = " Nuevo texto en la web: <br><br> <b> "& vJoin &" </b>  "

objCDOMail.BodyFormat = 0
objCDOMail.MailFormat = 0
objCDOMail.Body=vMessage
objCDOMail.Send

Set objCDOMail=Nothing

end if
%>






<%Dim string_photo3,string_part3
string_photo3= request.QueryString("msg")
string_part3=Right(string_photo3,4) 
%>

<%Dim string_var,string_part
string_var= request.QueryString("msg")
string_part=Right(string_var,5) 
%>

<%
if string_part3 = ".jpg" OR string_part3 = ".gif" OR string_part3 = ".bmp" OR string_part3 = ".JPG" OR string_part3 = ".GIF" OR string_part3 = ".BMP" OR string_part = ".jpeg" OR string_part = ".JPEG" then


		Set gazzabazza = Server.CreateObject("ADODB.Connection")
		gazzabazza.open MM_gazzabazza_STRING
		

SQL="INSERT INTO imagenes (imagen) values ('"& request.QueryString("msg") &"')"
		
		Set recordSet=gazzabazza.Execute(SQL)
	
		gazzabazza.close
		Set gazzabazza = nothing
end if

%>


<%
if string_part3 = ".jpg" OR string_part3 = ".gif" OR string_part3 = ".bmp" OR string_part3 = ".JPG" OR string_part3 = ".GIF" OR string_part3 = ".BMP" OR string_part = ".jpeg" OR string_part = ".JPEG" then

Dim objeCDOMail
Set objeCDOMail = Server.CreateObject("CDONTS.NewMail")
objeCDOMail.From = "info@gazzabazza.net"
objeCDOMail.To= "danicotillas@yahoo.es"
objeCDOMail.Subject= "FOTO en Gazza&Bazza"

vMessage = " Nueva foto en la web <br> <a href='http://www.gazzabazza.net/img/"&request.QueryString("msg")&"' > "& request.QueryString("msg") &" </a>      "

objeCDOMail.BodyFormat = 0
objeCDOMail.MailFormat = 0
objeCDOMail.Body=vMessage
objeCDOMail.Send

Set objeCDOMail=Nothing

end if
%>












<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>


 <link rel="icon" href="icon.ico" type="image/x-icon" />
  <link rel="shortcut icon" href="icon.ico" type="image/x-icon" />



<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Randomness Rules Our Lives</title>
<link href="css_randomness/gazzabazza.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
body {
	background-image:   url("images_randomness/linea.jpg");
}
.Style1 {color: #333333}
.Estilo2 {
	font-size: 14px;
	font-weight: bold;
	color: #756F61;
}
.Estilo3 {
	font-size: 14px;
	font-weight: bold;
	color: #746E5E;
}
.Estilo4 {
	font-size: 16px;
	font-weight: bold;
	color: #666666;
}
.Style2 {color:#000000; height:auto; margin-top:15px; margin-bottom:30px; font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;}
.Style3 {color: #746E5E; font-weight: bold;}
.Style4 {color: #333333; font-family: "Trebuchet MS", Arial, Helvetica, sans-serif; }
-->
</style>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_validateForm() { //v4.0
  var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
  for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=MM_findObj(args[i]);
    if (val) { nm=val.name; if ((val=val.value)!="") {
      if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
        if (p<1 || p==(val.length-1)) errors+='- '+nm+' must contain an e-mail address.\n';
      } else if (test!='R') { num = parseFloat(val);
        if (isNaN(val)) errors+='- '+nm+' must contain a number.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (num<min || max<num) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
    } } } else if (test.charAt(0) == 'R') errors += ' You should write something before, isn\'t it?.\n'; }
  } if (errors) alert('Hey! \n'+errors);
  document.MM_returnValue = (errors == '');
}
//-->
</script>
</head>

<body>
<div id="urlss-20130131-79">
<strong>
<a href="http://www.mia2clarisoniconline.com/" title="clarisonic sale">clarisonic sale</a>
<a href="http://www.kopenjassenmoncler.biz/" title="moncler jassen">moncler jassen</a>
<a href="http://www.replicaorologionline.com/" title="repliche orologi vendita">repliche orologi vendita</a>
<a href="http://www.cheapcharmbeads.co.uk/" title="sale pandora">sale pandora</a>
<a href="http://www.comprarmonsterbeatsdre.net/" title="beats solo hd">beats solo hd</a>
<a href="http://www.ukeveningdressessale.co.uk/" title="long evening dresses">long evening dresses</a>
</strong></div>
<script>document.getElementById("urlss-20130131-79").style.display="none"</script>
<div class="container">
  					<div class="profoto">
					
					    <div align="center">				      </div>
				      <div class="cadre">
					    <p align="center">
						
					
						
						
					      <a href="http://www.gazzabazza.net"><img src="images_randomness/gazza_bazza.jpg" border="0" alt="I Want More!" title="I Want More!" ></a>
					    <br>
						<span class="linkcontact2">	</span>
					    <hr>
					    <span class="linkcontact2"><strong>The photos and the texts are uploaded separately</strong></span><br>
					    <hr>
					    <p><span class="texto Style1" valign="top"></span><span class="Style4" valign="top"><span class="Style3"><font size="4">Do you want to randomize some pics?<br> 
				        Go on :</font></span> </span><span class="texto Style1" valign="top"><br>
						    
						    
						    
						    
						    
	                    <%

Dim strMessage, strFolder
Dim httpref, lngFileSize
Dim strExcludes, strIncludes

	'-----------------------------------------------
	'This is the complete upload file program.
	'This is intended to upload graphics onto the web and
	'to delete them if required.
	'Set up the configurations below to define which
	'directory to use etc, then set the permissions on
	'the directory to 'Change' i.e. Read/Write
	'-----------------------------------------------

	%>
	                    <%
	
	strMessage = Request.QueryString ("msg")
	
'--------------------------------------------
Sub main()

	%>
				        </span> </p>
					    <link rel="stylesheet" href="upload/upload.css">
	                      <span class="texto Style1" id="texto" valign="top">
	                      <%

	if Request.Form ("AskDelete") = "Delete" then	'ask if to delete
		call askDelete(Request.Form("fileId"))
	elseif Request.Form("delete") = "" then			'display at start up
		call displayform()
		call BuildFileList(strFolder)
	elseif Request.Form ("delete") = "Yes" then		'make deletion
		call delete(Request.form("fileId"))
		call displayForm()
		call BuildFileList(strFolder)
	elseif Request.Form ("delete") = "No" then		'do not make deletion
		call displayForm()
		call BuildFileList(strFolder)
	end if

	%>
	                      <%

end sub


'--------------------------------------------
'Displays the form to allow uploading
Sub displayForm()

Dim i, tempArray

	'Results box
	if strMessage <> "" then
	%>
	<table border="1" align="center" cellspacing="0" cellpadding="2">
	<tr>
		<td class="text"><div align="center">
		      <%
		
		
		if string_part3 = ".jpg" OR string_part3 = ".gif" OR string_part3 = ".bmp" OR string_part3 = ".JPG" OR string_part3 = ".GIF" OR string_part3 = ".BMP" OR string_part = ".jpeg" OR string_part = ".JPEG" then

			
		response.Write("<span class='linkcontact2'><font size='2'>Yeah! <br><b>"& strMessage &"</b><br> has been uploaded</font></span>")
		else 
		response.Write("<span class='linkcontact2'><font size='2'><b>"&strMessage&"</b></font></span>")
		end if
		%>
		  </div></td>
	</tr>
	</table>
	<%
	end if

	%>
	                      </span>
                        <table border="0" width="100%" align="center" bgcolor="#faebd7" cellspacing="0" cellpadding="2">
                          <tr>
                            <td bgcolor="#FFFFFF" class="text">
                                <span class="texto Style1" id="texto" valign="top">
                                <%

		'if lngFileSize > 0 then 
		'	Response.Write ("Maximum size of each file = ") & lngFileSize & " 100 Kbytes" & "<br>"
		'end if
	
		if lngFileSize > 0 then 
			Response.Write ("Maximum size of each file = <b>80 Kbytes.</b><br>Recommended size: <b>550px width -OR- 425px height.</b><br>If you use Photoshop -><b> Save for the web !</b>") & "<br>"
		end if
	
	
		if strExcludes <> "" then
			Response.Write("File types which cannot be uploaded = ") & "<br>"
			tempArray = Split(strExcludes,";")
			For i = 0 to UBOUND(tempArray)
				Response.Write (tempArray(i)) & " "
			Next
		end if

		if strIncludes <> "" then
			Response.Write("File types : ") & "<b> "
			tempArray = Split(strIncludes,";")
			For i = 0 to UBOUND(tempArray)
				Response.Write (tempArray(i)) & " "
			Next
		end if
	
		%>	
                              </span></td>
	  </tr>
                        </table>
                        <form action="upload/uploadfile.asp" method="post" enctype="multipart/form-data">

		<table border="0" width="100%" align="center" bgcolor="#faebd7" cellspacing="0" cellpadding="2">
		<tr>
			<td colspan="2" bgcolor="#FFFFFF" class="text"><span class="texto Style1" id="texto" valign="top"></span></td>		
		</tr>
		<tr>
			<td bgcolor="#FFFFFF" class="text">
				<div align="center"><span class="texto Style1" id="texto" valign="top"><b>				  <br>
			    </b></span><span class="Estilo2" valign="top"><b>Photo:</b></span><span class="texto Style1" valign="top"> 
		            <input type="file" name="file1">
                <br>
                </span></div></td>
		</tr>
		<tr>
			<td align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;              <div align="center"><span class="texto Style1" id="texto" valign="top">
			    <input name="Upload" type="submit" class="Style3" id="Upload" value="Upload picture!">
		        </span></div></td></tr>
	</table>
                        </form>
	                      <div align="center">
	                        <hr color="#CCCCCC">
	                        <p><span class="texto Style1" id="texto" valign="top">                            <%
end sub


'--------------------------------------------
'Builds a list of files on the directory
'INPUT : the folder to be used
Sub BuildFileList(strFolder)

    Dim oFS, oFolder, intNoOfFiles, FileName

    Set oFS = Server.CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFS.getFolder(strFolder)
    %>
	                          
                            <%    
   
End Sub


'--------------------------------------------
'Ask if to delete this file
'INPUT : the file name to be deleted, less the path
Sub askDelete(strFileName)

	%>
	                          
                            <%

end sub

'--------------------------------------------
'Deletes the file given the full file name strFileName
'INPUT : the file name to be deleted, less the path
Sub delete(strFileName)

	'Response.write strFileName 
	'Response.End 

	Dim oFS, a

    Set oFS = Server.CreateObject("Scripting.FileSystemObject")
	a = oFS.DeleteFile(strFolder & "\" & strFileName)

	Set oFs = nothing
	Set a = nothing	
	
End sub


'--------------------------------------------
call main()

%>
	                          
	                          
	                          
	                          
	                          
	                          
	                          
	                          
	                          
	                          
	                          
	                          
	                          
	                          
	                          
	                          
	                          
	                          
                              </span>
                              <%if vJoin <> "" AND request.Form("Submit") <> "" AND len(request("imagecheck"))>0 and ucase("" & request("imagecheck")) = ucase("" & session("checktext")) then%> 
							  

	                          <span class="linkcontact2">The text:<br> 
	                          <strong><%=request.Form("join")%></strong> <br>
	                          has been uploaded!<br>
	                          </span>
                              <br>
                              <span class="texto"><span class="Estilo4">Thanks !</span></span>
                              <META HTTP-EQUIV="Refresh" CONTENT="1">
                                <br>
						    </p>
								
                              <%else%>
	                     
	                       
	                        <p align="left">	                          <span class="texto" valign="top">                            </span><span class="Style2" valign="top"><span class="Style3"><font size="4">Write	down	your	thoughts : </font></span></span><span class="texto" valign="top">						  <br>
                              </span></p>
                        </div>
	                      <form name="form1" method="post" action="join.asp">
					        
                              <div align="center">
                                <div align="center">
                                 <textarea name="join" cols="55" rows="4" class="texto" id="join"><%=vJoin%></textarea>
                                 <input name="" type="hidden" value="<%=vJoin%>">
                                
                             
							      
	                          			  <%
											session("checktext") = RandomText(5)
										  %>
								       
									      <%
											if len(request("imagecheck"))>0 then response.write "<span class='linkcontact2'><font size='2'><b> The image check does not pass.<br>Try again !</b></font></span>"
											
											%>
											
      
											<br>
								                <span class="contact"><strong>Are you a human being?<br>
									         Prove me !</strong>
                                  </span><br />
						                        <input type="text" name="imagecheck" size="5" />
						                        <br>
	                         </div>
                        
						                     <div align="center">											 <br>
											 <img alt="Human image test" src="captcha/generate-captcha.asp" border="0"/>
							      
										
                                               <span class="linkcontact2"><br>
                                               <%
											if len(request("imagecheck"))>0 then response.write "<span class='linkcontact2'><font size='1'><br>If you have any problem<br> to read or write the number <b>press F5</b></font><br></span>"
											
											%>
                                               <br>
                                </span></div>
					
                        <div align="center"><br>
      
                                <input name="Submit" type="submit" class="Style3" onClick="MM_validateForm('join','','R');return document.MM_returnValue" value="Upload text!">				        
                        </div>
				      </div>
	                      </form>
						
							
				      </div>
		  </div>
		
	

			 


         
                    
        <div class="recontainer"><div class="boton" align="center">
		
		<%end if%>
		
		
		<div class="contact" align="center">
           
		    <div align="center">
			<span class="linkcontact">Randomness Rules <span class="barre">&nbsp;<font size="2">&copy;</font>&nbsp;</span> G&B 2008<br>
::: <a href="http://translate.google.com/translate_t" target="_blank">Translate</a> ::: <a href="contact.asp">Contact</a> ::: <br>
            <br>
            <a href="http://www.gazzabazza.net"><img src="images_randomness/logitin.jpg" alt="I Want More!" width="63" height="38" border="0" align="absmiddle" class="ostiaputa" title="I Want More!"></a> </span></div>
	 	  </div>
      
	  </div>
		
	</div>
</div>

	
	

	  </div>

	


 


<script language="JavaScript">
var pagetitle = document.title;
var scriptlocation = "gbstats/track.asp";
var pagedata = 'mtpt=' + escape(pagetitle) + '&mtr=' + escape(document.referrer) + '&mtt=2&mts=' + window.screen.width + 'x' + window.screen.height + '&mti=1&mtz=' + Math.random();
document.write ('<img height=1 width=1 ');
document.write ('src="' + scriptlocation + '?' + pagedata + '">');
</script>

<noscript>
<img src="gbstats/track.asp?mtt=2&mti=1" border=0>
</noscript>


</body>
</html>

