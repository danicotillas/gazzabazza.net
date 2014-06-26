<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/gazzabazza.asp" -->
<%

if request.form("Name") <> "" then
	vName = request.form("Name")
	vName = replace(vName,"'","&#39;")
else
	vName = ""
end if


if request.form("Subject") <> "" then
	vSubject = request.form("Subject")
	vSubject = replace(vSubject,"'","&#39;")
else
	vSubject = ""
end if

if request.form("Message") <> "" then
	vMessages = request.form("Message")
	vMessages = Replace(vMessages,chr(10),"<br>")
	vMessages  = replace(vMessages,"'","&#39;")
	
else
	vMessages  = ""
end if
%>


<%
if request.form("Enviar") <> "" then

	Dim objCDOMail
	Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
	objCDOMail.From = request.Form("email")
	objCDOMail.To = "danicotillas@gmail.com"
	objCDOMail.Subject= "Message from Gazza&Bazza -  " & vSubject
	vMessage =  vName & " has contacted with G&B to say: " &"<br><br>"&  vMessages
	objCDOMail.BodyFormat = 0
	objCDOMail.MailFormat = 0
	objCDOMail.Body=vMessage
	objCDOMail.Send
	
	Set objCDOMail=Nothing

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
<script type="text/JavaScript">
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
        if (p<1 || p==(val.length-1)) errors+='- '+nm+' debe ser un email para que pueda contestarte.\n';
      } else if (test!='R') { num = parseFloat(val);
        if (isNaN(val)) errors+='- '+nm+' must contain a number.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (num<min || max<num) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
    } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' debe ser rellenado.\n'; }
  } if (errors) alert('Pero...¿Cómo es posible?! Por favor, mira a ver:\n'+errors);
  document.MM_returnValue = (errors == '');
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>

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
.Estilo5 {font-family: Verdana, Arial, Helvetica, sans-serif; color: #736e5f; }
.Style2 {color: #756F61; font-weight: bold;}
.Style3 {color:#736e5f; text-align:center; font-family: Verdana, Arial, Helvetica, sans-serif;}
-->
</style>

</head>

<body>

<div class="container">
  					<div class="profoto">
					
				      <div class="cadre" style="width:300px;">
				        <p align="center"><a href="http://www.gazzabazza.net"><img src="images_randomness/gazza_bazza.jpg" border="0" alt="¡Quiero m&aacute;s!" title="¡Quiero m&aacute;s!" ></a>                        
					    <p align="center" class="contact"><strong>					    <font face="Georgia, Times New Roman, Times, serif"><font face="Verdana, Arial, Helvetica, sans-serif">Dime qu&eacute; te pareci&oacute; lo que ves<br> 
				      o preguntame lo que te parezca</font></font></strong> 
					    <div align="center">
				            <%
  if vName <> "" AND request.form("email") <> "" AND vSubject <> ""AND vMessages <> ""then
   %>
     
                <span class="boton"><br>
                               <br>
                               <br>
                               Super! Tu mensaje ya parti&oacute;!<br>
                               Gracias!</span><br>
                               <br />
                                 <span class="linkcontact"> <a href="http://www.gazzabazza.net" class="linkcontact2">Volver!</a> <br>
                                 <br> <br> <br> <br>
                </span>
        </div>
  <%
  else
  %>
  <form action="contact.asp" method="post" name="form1" id="form1" onSubmit="MM_validateForm('Name','','R','email','','RisEmail','Subject','','R','Message','','R');return document.MM_returnValue" >
<table width="85%" height="100%" border="0" align="center">
                            
                            
                            <tr>
                              <td width="31" height="31">&nbsp;</td>
                              <td width="120" valign="top" class="contact"><strong>&iquest;Qui&eacute;n?</strong> </td>
                              <td width="225" valign="top"><input name="Name" type="text" id="Name" /></td>
                            </tr>
                            <tr>
                              <td>&nbsp;</td>
                              <td valign="top"><span class="contact"><strong>&iquest;D&oacute;nde?</strong></span><span class="Estilo5"><br />
<br />
                                                            </span></td>
                              <td valign="top"><input name="email" type="text" id="email" value="Email por favor" />                              </td>
                            </tr>
                            <tr>
                              <td height="32">&nbsp;</td>
                              <td valign="top"><span class="contact"><strong>&iquest;Porqu&eacute;?</strong> </span><span class="Estilo5"><br />
                              </span></td>
                              <td valign="top"><input name="Subject" type="text" id="Subject" /></td>
                            </tr>
                            <tr>
                              <td height="103">&nbsp;</td>
                              <td valign="top" class="contact"><strong>&iquest;Qu&eacute;?</strong></td>
                              <td valign="top"><textarea name="Message" rows="10" id="Message" ></textarea>                              </td>
                            </tr>
                            <tr>
                              <td>&nbsp;</td>
                              <td>&nbsp;</td>
                              <td><label>
                                <input name="Enviar" type="submit" id="Enviar" value="Dale!" />
                              </label></td>
                            </tr>
        </table>
					    </form>
						



						
							  
				 </div>
				 
				 
				 
				 
				 
		  </div>
		
	
		
			 
      
	  </div>
	  
	  
<div class="recontainer"><div class="boton" align="center">
    <%end if%>
  <div class="contact" >
           
		    <div align="center"><span class="linkcontact"><strong><a href="join_images.asp">SUBIR IM&Aacute;GENES</a></strong>&nbsp; //&nbsp; <strong><a href="join_text.asp">SUBIR TEXTO</a></strong> <br>
<br>
            </span><span class="linkcontact">:::<strong><a href="about.asp">&nbsp; Sobre G&B</a></strong>&nbsp; ::: <a href="http://translate.google.com/translate_t#" target="_blank">Traductor</a> ::: <a href="contact.asp">Contacta</a> ::: </span><br>
El Azar Manda <span class="barre">&nbsp;<font size="2">&copy;</font>&nbsp;</span><span class="linkcontact">2009</span></div>
		    <br>
<a href="http://creativecommons.org/licenses/by-nc-sa/3.0/es/" target="_blank" rel="license"><img src="http://i.creativecommons.org/l/by-nc-sa/3.0/es/80x15.png" alt="Creative Commons License" border="0" style="border-width:0" /></a><br />
                                    <span xmlns:dc="http://purl.org/dc/elements/1.1/" href="http://purl.org/dc/dcmitype/InteractiveResource" property="dc:title" rel="dc:type"></span>	
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
