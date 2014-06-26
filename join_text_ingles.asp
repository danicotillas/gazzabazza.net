<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!-- #include file = "upload/config.asp" -->
	<!--#INCLUDE FILE="captcha/_captcha.asp"-->
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

<div class="container">
  					<div class="profoto">
					
					    <div align="center">				      </div>
				      <div class="cadre">
					    <p align="center">
						
					
						
						
					      <a href="http://www.gazzabazza.net"><img src="images_randomness/gazza_bazza.jpg" border="0" alt="I Want More!" title="I Want More!" ></a>
					    <br>
						<span class="linkcontact2">	</span>
					    <hr color="#999999">
					    <p><span class="texto Style1" valign="top"></span></p>
					    <form action="upload/uploadfile.asp" method="post" enctype="multipart/form-data">
					    </form>
                        <div align="center">
	                       
	                        <p><span class="texto Style1" id="texto" valign="top">                   
	                          
	                          
	                          
	                          
	                          
	                          
	                          
	                          
	                          
	                          
	                          
	                          
	                          
	                          
	                          
	                          
	                          
	                          
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
	                      <form name="form1" method="post" action="join_text.asp">
					        
                              <div align="center">
                                <div align="center">
                                 <textarea name="join" cols="45" rows="4" wrap="virtual" class="texto" id="join"><%=vJoin%></textarea>
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
			<span class="linkcontact"><strong><a href="join_images.asp">UPLOAD IMAGES</a></strong>&nbsp; //&nbsp; <strong><a href="join_text.asp">UPLOAD TEXT</a></strong> <br>
<br>
			<br>
			Randomness Rules <span class="barre">&nbsp;<font size="2">&copy;</font>&nbsp;</span> G&amp;B 2009<br>
			:::<strong><a href="about.asp">&nbsp; About</a></strong>&nbsp; ::: <a href="http://translate.google.com/translate_t#" target="_blank">Translate</a> ::: <a href="contact.asp">Contact</a> :::<br>
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
