<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/gazzabazza.asp" -->




<%
Function RandomNumberImagen(inicio, fin)
Randomize time
RandomNumberImagen = Int((fin - inicio + 1) * Rnd + inicio)
End Function
%> 


<%
Function RandomNumberTexto(inicio, fin)
Randomize
RandomNumberTexto = Int(((fin - inicio + 1) * Rnd) + inicio)
End Function
%>



<%
Dim rsTexto_var
Dim rsTexto_var_numRows

Set rsTexto_var = Server.CreateObject("ADODB.Recordset")
rsTexto_var.ActiveConnection = MM_gazzabazza_STRING
rsTexto_var.Source = "SELECT MAX(id) FROM texto"
rsTexto_var.CursorType = 0
rsTexto_var.CursorLocation = 2
rsTexto_var.LockType = 1
rsTexto_var.Open()

rsTexto_var_numRows = 0
%>

<%
Dim rsImages_var
Dim rsImages_var_numRows

Set rsImages_var = Server.CreateObject("ADODB.Recordset")
rsImages_var.ActiveConnection = MM_gazzabazza_STRING
rsImages_var.Source = "SELECT MAX(id) FROM imagenes"
rsImages_var.CursorType = 0
rsImages_var.CursorLocation = 2
rsImages_var.LockType = 1
rsImages_var.Open()

rsImages_var_numRows = 0
%>

<%
Dim rImagen
Dim rImagen_numRows

Set rImagen = Server.CreateObject("ADODB.Recordset")
rImagen.ActiveConnection = MM_gazzabazza_STRING
rImagen.Source = "SELECT imagen FROM imagenes WHERE id = "& RandomNumberImagen(1,(rsImages_var.Fields.Item("Expr1000").Value)) &" "
rImagen.CursorType = 0
rImagen.CursorLocation = 2
rImagen.LockType = 1
rImagen.Open()

rImagen_numRows = 0
%>
<%
Dim rTexto
Dim rTexto_numRows

Set rTexto = Server.CreateObject("ADODB.Recordset")
rTexto.ActiveConnection = MM_gazzabazza_STRING
rTexto.Source = "SELECT * FROM texto WHERE id = "& RandomNumberTexto(1,(rsTexto_var.Fields.Item("Expr1000").Value) )&""
rTexto.CursorType = 0
rTexto.CursorLocation = 2
rTexto.LockType = 1
rTexto.Open()

rTexto_numRows = 0
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>


 <link rel="icon" href="icon.ico" type="image/x-icon" />
  <link rel="shortcut icon" href="icon.ico" type="image/x-icon" />


<meta http-equiv="Page-Exit" content="blendTrans(Duration=0.7)">
<META HTTP-EQUIV="Refresh" CONTENT="5">


<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="description" content="La raz&oacute;n nos limita, el azar es INFINITO. Es &eacute;l quien nos descubrir&aacute; los sentidos ocultos, los sabores desconocidos, los placeres an&oacute;nimos que esconden las im&aacute;genes cotidianas. Sus significados son m&uacute;ltiples y extra&ntilde;os, y es aqu&iacute;, cuando definitivamente, encontrar&aacute;s la llave." />
<meta name="verify-v1" content="x8SGvOjKSa0+DPztRByBbbPpVCD4J7mANcCQIn/48SY=" />

<title>Gazza &amp; Bazza ::: Randomness Rules Our Lives</title>
<link href="css_randomness/gazzabazza3.css" rel="stylesheet" type="text/css">

<script language="JavaScript" type="text/JavaScript">
<!--



function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}
//-->
</script>
</head>

<body onLoad="MM_preloadImages('images_randomness/logitin.jpg')">



<form method="post" action="">

	
	
			
		<div class="container">
		
				<div class="profoto">

						<div align="center"><img src="img/<%=(rImagen.Fields.Item("imagen").Value)%>" align="middle" />
						</div>
						
				    	<div class="cadre">
						
								<span class="texto" id="texto" valign="top">
								<br>
								<%response.Write((rTexto.Fields.Item("ideas").Value))%>
								<br>
								</span>
								
						</div>
						
				</div>
         

                  
				  
				<div class="recontainer">
							
						<div class="boton" align="center">
		 							
								<div class="contact" align="center">
           							
								  <div align="center">
										
           								<p><span class="linkcontact"><img src="images_randomness/logitin.jpg" alt="I Want More!" width="63" height="38" border="0" align="absmiddle" class="ostiaputa" title="I Want More!"></span>            							  <br>
       								      <span class="linkcontact">       								      </span>
       								      <strong><span class="barre"><font size="2">&copy;</font></span><span class="linkcontact"></span> - El Azar Manda en Nuestras Vidas</strong><span class="barre"><br>
       								      </span><span class="linkcontact"><br>
</span>					                </p>
   								  </div>
										
								</div>
									
						</div>
	  
	                  <!-- <div align="center"><font color="#FFFFFF" style="font-size:3px; font:Verdana, Arial, Helvetica, sans-serif; ">texto:<%=(rsTexto_var.Fields.Item("Expr1000").Value)%> | photos: <%=(rsImages_var.Fields.Item("Expr1000").Value)%></font> -->
		  
				</div>
					
		</div>
		

</form>




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
rImagen.Close()
Set rImagen = Nothing
%>
<%
rTexto.Close()
Set rTexto = Nothing
%>
<%
rsImages_var.Close()
Set rsImages_var = Nothing
%>
<%
rsTexto_var.Close()
Set rsTexto_var = Nothing
%>
