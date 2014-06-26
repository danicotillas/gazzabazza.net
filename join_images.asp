<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!-- #include file = "upload/config.asp" -->
	<!--#INCLUDE FILE="captcha/_captcha.asp"-->
	<!-- #include file = "Connections/gazzabazza.asp" -->
	
	




<%Dim string_photo3,string_part3
string_photo3= request.QueryString("msg")
string_part3=Right(string_photo3,4) 
%>

<%Dim string_var,string_part
string_var= request.QueryString("msg")
string_part=Right(string_var,5) 
%>

<%

Set gazzabazza2 = Server.CreateObject("ADODB.Connection")
gazzabazza2.open MM_gazzabazza_STRING
SQL= "SELECT imagen FROM imagenes WHERE imagen='"&string_var&"'"
Set gazzabazza2 = gazzabazza2.Execute(SQL)

%>




<%
if gazzabazza2.EOF AND string_part3 = ".jpg" OR string_part3 = ".gif" OR string_part3 = ".bmp" OR string_part3 = ".JPG" OR string_part3 = ".GIF" OR string_part3 = ".BMP" OR string_part = ".jpeg" OR string_part = ".JPEG" then


		Set gazzabazza = Server.CreateObject("ADODB.Connection")
		gazzabazza.open MM_gazzabazza_STRING
		

SQL="INSERT INTO imagenes (imagen) values ('"& request.QueryString("msg") &"')"
		
		Set recordSet=gazzabazza.Execute(SQL)
	
		gazzabazza.close
		Set gazzabazza = nothing
end if

%>


<%
if gazzabazza2.EOF AND string_part3 = ".jpg" OR string_part3 = ".gif" OR string_part3 = ".bmp" OR string_part3 = ".JPG" OR string_part3 = ".GIF" OR string_part3 = ".BMP" OR string_part = ".jpeg" OR string_part = ".JPEG" then

Dim objeCDOMail
Set objeCDOMail = Server.CreateObject("CDONTS.NewMail")
objeCDOMail.From = "info@gazzabazza.net"
objeCDOMail.To= "danicotillas@gmail.es"
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
<title>El Azar Manda en Nuestras Vidas</title>
<link href="css_randomness/gazzabazza.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
body {
	background-image:   url("images_randomness/linea.jpg");
}
.Style1 {
font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
color: #333333;
font-size: 14px;
}
.Estilo2 {
font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
	font-size: 14px;
	font-weight: bold;
	color: #756F61;
}
.Estilo3 {
font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
	font-size: 12px;
	font-weight: bold;
	color: #746E5E;
}
.Estilo4 {
font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
	font-size: 10px;
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
					
					    <div class="cadre">
					    <p align="center">
						
					
						
						
					      <a href="http://www.gazzabazza.net"><img src="images_randomness/gazza_bazza.jpg" border="0" alt="¡Quiero m¨¢s!" title="¡Quiero m¨¢s!" ></a>
					    <br>
						<span class="linkcontact2">	</span>
					    <hr color="#999999">
					    <p><span class="Style4" valign="top"><span class="Style3"><font size="4">				        Quieres randomizar algunas fotos?<br> 
				        Dale :</font></span> </span><span class="texto Style1" valign="top"><br>
						    
						    
						    
						    
						    

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
  
  
                                 <%
			 If NOT gazzabazza2.EOF then
%>
			 <p class="linkcontact2"><strong>Vaya!<br>
			 </strong>
			   Ya existe una imagen con ese nombre.<br>
			   <br>
			   Por favor, <strong>c¨¢mbiale el nombre a tu foto</strong><br>
			   y vuelve a subirla!<br>
			   <br>
			   <strong>Gracias!</strong></p>
			
           
 <%else%> 
  
  
  
	<table border="1" align="center" cellspacing="0" cellpadding="2">
	<tr>
		<td class="text"><div align="center">
		      <%
		
		
		if string_part3 = ".jpg" OR string_part3 = ".gif" OR string_part3 = ".bmp" OR string_part3 = ".JPG" OR string_part3 = ".GIF" OR string_part3 = ".BMP" OR string_part = ".jpeg" OR string_part = ".JPEG" then

			
		response.Write("<span class='linkcontact2'><font size='2'>Si! <br><b>"& strMessage &"</b><br> ha sido subida con éxito!</font></span>")
		else 
		response.Write("<span class='linkcontact2'><font size='2'><b>"&strMessage&"</b></font></span>")
		end if
		%>
		  </div></td>
	</tr>
	</table>
	<br>
	<%
	end if
	%>
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
		'	Response.Write ("Maximum size of each file = ") & lngFileSize & " 150 Kbytes" & "<br>"
		'end if
	
		if lngFileSize > 0 then 
			Response.Write ("Máximo tamaño de archivo = <b>250 Kbytes.</b><br><br><font size='1'> TAMAÑO RECOMENDADO:</font> <b>800px ancho &#65533; 650px alto (72ppp)</b><br>Si utilizas Photoshop -><b> Guardar para WEB!</b>") & "<br>"
		end if
	
	
		if strExcludes <> "" then
			Response.Write("<font size='1'>Tipo de archivo que no puede ser subido = </font>") & "<br>"
			tempArray = Split(strExcludes,";")
			For i = 0 to UBOUND(tempArray)
				Response.Write (tempArray(i)) & " "
			Next
		end if

		if strIncludes <> "" then
			Response.Write("<font size='1'>Tipo de archivo: </font>") & "<b> "
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
			    </b></span><span class="Estilo2" valign="top"><b>Imagen/Foto:</b></span><span class="texto Style1" valign="top"> 
		            <input type="file" name="file1">
                <br>
                </span></div></td>
		</tr>
		<tr>
			<td align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;              <div align="center"><span class="texto Style1" id="texto" valign="top">
			    <input name="Upload" type="submit" class="Style3" id="Upload" value="Sube tu foto!">
		        </span></div></td></tr>
	</table>
                        </form>
                        <div align="center">
	                        
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
	                          
	                          
	                          
	                          
	                          
	                          
	                          
	                          
	                          
	                          
	                          
	                          
	                          
	                          
	                          
	                          
	                          
	                          
                          </span></p>
                        </div>
	                      <form name="form1" method="post" action="join_images.asp">
	                      </form>
						
							
				      </div>
		  </div>
		
	

			 


         
                    
        <div class="recontainer"><div class="boton" align="center">
		
	
		
		
		<div class="contact" align="center">
           
		    <div align="center">
			<span class="linkcontact"><strong><a href="join_images.asp">SUBIR IM&Aacute;GENES</a></strong>&nbsp; //&nbsp; <strong><a href="join_text.asp">SUBIR TEXTO</a></strong> <br>
			<br>
			:::<strong><a href="about.asp">&nbsp; Sobre G&B&nbsp;</a></strong> ::: <a href="http://www.twitter.com/danicotillas" target="_blank">@danicotillas</a><a href="http://translate.google.com/translate_t#" target="_blank"></a> ::: <a href="contact.asp">Contacta</a> :::<br>
			</span><br>
<a href="http://creativecommons.org/licenses/by-nc-sa/3.0/es/" target="_blank" rel="license"><img src="http://i.creativecommons.org/l/by-nc-sa/3.0/es/80x15.png" alt="Creative Commons License" border="0" style="border-width:0" /></a><br />
                                    <span xmlns:dc="http://purl.org/dc/elements/1.1/" href="http://purl.org/dc/dcmitype/InteractiveResource" property="dc:title" rel="dc:type"></span>	
</div>
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

<%
		gazzabazza2.close
		Set gazzabazza2 = nothing
%>
