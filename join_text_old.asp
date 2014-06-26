<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!-- #include file = "upload/config.asp" -->
	<!-- #include file = "Connections/gazzabazza.asp" -->
	
	




<%
if request.Form("join") <> "" then
	vJoin = request.Form("join")
	vJoin = replace(vJoin,"'","&#39;")
	vJoin = replace(vJoin,"é","&eacute;")
	vJoin = replace(vJoin,"á","&aacute;")
	vJoin = replace(vJoin,"í","&iacute;")
	vJoin = replace(vJoin,"ó","&oacute;")
	vJoin = replace(vJoin,"ú","&uacute;")
	vJoin = replace(vJoin,"ñ","&ntilde;")
	vJoin = replace(vJoin,"à","&agrave;")
	vJoin = replace(vJoin,"¿","&iquest;")
	vJoin = replace(vJoin,chr(10),"<br />")
	
else
	vJoin = ""
end if
%>


<%

Set gazzabazza2 = Server.CreateObject("ADODB.Connection")
gazzabazza2.open MM_gazzabazza_STRING
SQL= "SELECT ideas FROM texto WHERE ideas='"&vJoin&"'"
Set gazzabazza2 = gazzabazza2.Execute(SQL)

%>



<%
if vJoin <> "" AND request.Form("Submit") <> "" AND gazzabazza2.EOF then

		Set gazzabazza = Server.CreateObject("ADODB.Connection")
		gazzabazza.open MM_gazzabazza_STRING
		

SQL="INSERT INTO texto (ideas) values ('"& vJoin &"')"
		
		Set recordSet=gazzabazza.Execute(SQL)
	
		gazzabazza.close
		Set gazzabazza = nothing
end if

%>


<%
if vJoin <> "" AND request.Form("Submit") <> "" AND gazzabazza2.EOF then

Dim objCDOMail
Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
objCDOMail.From = "info@gazzabazza.net"
objCDOMail.To= "danicotillas@gmail.com"
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



<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>

<script type="text/javascript" src="http://use.typekit.com/zgr2ihx.js"></script>
<script type="text/javascript">try{Typekit.load();}catch(e){}</script>



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
    } } } else if (test.charAt(0) == 'R') errors += ' ¿Qué? ¿No vas a escribir nada?\n'; }
  } if (errors) alert(errors);
  document.MM_returnValue = (errors == '');
}
//-->
</script>
</head>

<body>

<div class="container">
  					<div class="profoto">
					
				      <div class="cadre">
					    <p align="center">
						
					
						
						
					      <a href="http://www.gazzabazza.net"><img src="images_randomness/gazza_bazza.jpg" border="0" alt="¡Quiero m&aacute;s!" title="¡Quiero m&aacute;s!" ></a>
					    <br>
						<span class="linkcontact2">	</span>
					    <hr color="#999999">
					    <p><span class="texto Style1" valign="top"></span></p>
					    <form action="upload/uploadfile.asp" method="post" enctype="multipart/form-data">
					    </form>
                        <div align="center">
	                       
	                        <p><span class="texto Style1" id="texto" valign="top">                   
	                               
	                          
	                          
                              </span>
                              
                               <%
			 If NOT gazzabazza2.EOF then
%>
			 <p class="linkcontact2"><strong>¿Te creías original?<br>
			   <br>
			 </strong>Pues no lo fuiste esta vez!<br>
			   Este texto ya se le ocurrió a alguien antes.<br>
			   Dale, escribe otro pues!<br>
			   <strong>Gracias!</strong></p>
			 <p class="linkcontact2"><strong>Espera un segundo y puedes volver a intentarlo!</strong>
			   
			   </p>
			 <META HTTP-EQUIV="Refresh" CONTENT="4">

           
 <%else%>          
           
           <%if vJoin <> "" AND request.Form("Submit") <> "" AND gazzabazza2.EOF then%> 
							  

	                          <span class="linkcontact2">El texto:<br> 
	                          <strong><%=request.Form("join")%></strong> 

	                          <br>
	                          ha sido incluido con &eacute;xito!<br>
	                          </span>
                              <br>
                              <span class="texto"><span class="Estilo4">Gracias !</span></span><br>
<br>

                               <p class="linkcontact2"><strong>Espera un segundo y contin&uacute;a!</strong>			   </p>
                              
                               <META HTTP-EQUIV="Refresh" CONTENT="4">
						  
								
                              <%else%>
	                     
	                       
	                        <div align="left">	                                                      <span class="Style2" valign="top"><span class="Style3"><font size="4">Escribe tus pensamientos : </font></span></span><span class="texto" valign="top">						  
                              </span></div>
                        </div>
	                      <form name="form1" method="post" action="join_text.asp">
					        
                              <div align="center">
                                <div align="center">
                                 <textarea name="join" cols="30" rows="4" wrap="virtual" class="texto" id="join"><%=vJoin%></textarea>
                                 <input name="" type="hidden" value="<%=vJoin%>">
                                
                
	                         </div>
                        
						                     <div align="center">											 <br>
											 
							      
										
                                             </div>
					
                        <div align="center">
      
                                <input name="Submit" type="submit" class="Style3" onClick="MM_validateForm('join','','R');return document.MM_returnValue" value="Subir el texto!">				        
                        </div>
				      </div>
	                      </form>
						
							
				      </div>
		  </div>
		
            <div class="recontainer"><div class="boton" align="center">
		
 
		
		<div class="contact" align="center">
           
		    <div align="center">
			<span class="linkcontact"><strong><a href="join_images.asp">SUBIR IM&Aacute;GENES</a></strong>&nbsp; //&nbsp; <strong><a href="join_text.asp">SUBIR TEXTO</a></strong> <br>
			<br>
			:::<strong><a href="about.asp">&nbsp; Sobre G&amp;B</a></strong>&nbsp; ::: <a href="http://www.twitter.com/danicotillas" target="_blank">@danicotillas</a><a href="http://translate.google.com/translate_t#" target="_blank"></a> ::: <a href="contact.asp">Contacta</a> :::<br>
			</span><br>
<a href="http://creativecommons.org/licenses/by-nc-sa/3.0/es/" target="_blank" rel="license"><img src="http://i.creativecommons.org/l/by-nc-sa/3.0/es/80x15.png" alt="Creative Commons License" border="0" style="border-width:0" /></a><br />
                                    <span xmlns:dc="http://purl.org/dc/elements/1.1/" href="http://purl.org/dc/dcmitype/InteractiveResource" property="dc:title" rel="dc:type"></span>	<%
end if
%>
          
        </div>
	 	  </div>
      			
	  </div>

	</div>
	
	
	<%end if%>
	

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