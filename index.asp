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
rsTexto_var.Source = "SELECT MAX(Id) FROM texto"
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

<!--<script type="text/javascript" src="http://use.typekit.com/zgr2ihx.js"></script>
<script type="text/javascript">try{Typekit.load();}catch(e){}</script>-->


 <link rel="icon" href="icon.ico" type="image/x-icon" />
  <link rel="shortcut icon" href="icon.ico" type="image/x-icon" />

<meta name="google-site-verification" content="trmgbMmR4Ya7HnK786GGNZLhZskkZD9jCkaij4uYLqc" />
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="description" content="La raz&oacute;n nos limita, el azar es INFINITO. Es &eacute;l quien nos descubrir&aacute; los sentidos ocultos, los sabores desconocidos, los placeres an&oacute;nimos que esconden las im&aacute;genes cotidianas. Sus significados son m&uacute;ltiples y extra&ntilde;os, y es aqu&iacute;, cuando definitivamente, encontrar&aacute;s la llave." />
<meta name="verify-v1" content="x8SGvOjKSa0+DPztRByBbbPpVCD4J7mANcCQIn/48SY=" />

<title><%response.Write((rTexto.Fields.Item("ideas").Value))%> - Gazza & Bazza: El Azar Manda en Nuestras Vidas</title>
<link href="css_randomness/gazzabazza.css" rel="stylesheet" type="text/css">

<script>
  (function(i,s,o,g,r,a,m){i['GoogleAnalyticsObject']=r;i[r]=i[r]||function(){
  (i[r].q=i[r].q||[]).push(arguments)},i[r].l=1*new Date();a=s.createElement(o),
  m=s.getElementsByTagName(o)[0];a.async=1;a.src=g;m.parentNode.insertBefore(a,m)
  })(window,document,'script','//www.google-analytics.com/analytics.js','ga');

  ga('create', 'UA-51390850-1', 'gazzabazza.net');
  ga('send', 'pageview');

</script>

<script language="JavaScript" type="text/JavaScript">
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
</head>

<body>



<form method="post" action="">

	

			
		<div class="container">
		
<div class="profoto">


<img src="img/<%=(rImagen.Fields.Item("imagen").Value)%>" style="max-width:800px; max-height:425px;" /><br>
	
                                <div class="texto">
								  <%response.Write((rTexto.Fields.Item("ideas").Value))%>
								</div>
    
              
								
            
            		
			</div>
         

                  
				  
				<div class="recontainer">
							
						<div class="boton" align="center">
		 							
								<div class="contact" align="center">
           							
								  <div align="center">
										
           								<p>
            							  <input type="submit" name="Submit" value="Quiero m&aacute;s!" class="ostiaputa">
            							  <br>
            							  <br>
       								      <span class="linkcontact">       								 <strong><a href="join_images.asp">SUBIR IM&Aacute;GENES</a></strong>&nbsp; //&nbsp; <strong><a href="join_text.asp">SUBIR TEXTO</a></strong><br>
       								      <br>
       								      </span><span class="linkcontact">:::<strong><a href="about.asp">&nbsp;Sobre G&amp;B</a></strong>&nbsp; ::: <a href="http://www.twitter.com/danicotillas" target="_blank">@danicotillas</a> ::: <a href="contact.asp">Contacta</a> ::: </span><br>
       								             								      <br>
                                    <a href="http://creativecommons.org/licenses/by-nc-sa/3.0/es/" target="_blank" rel="license"><img src="http://i.creativecommons.org/l/by-nc-sa/3.0/es/80x15.png" alt="Creative Commons License" border="0" style="border-width:0" /></a><br />
                                    <span xmlns:dc="http://purl.org/dc/elements/1.1/" href="http://purl.org/dc/dcmitype/InteractiveResource" property="dc:title" rel="dc:type"></span>					                </p>
   								  </div>
										
								</div>
									
						</div>
	  
	                  <!-- <div align="center"><font color="#FFFFFF" style="font-size:3px; font:Verdana, Arial, Helvetica, sans-serif; ">texto:<%=(rsTexto_var.Fields.Item("Expr1000").Value)%> | photos: <%=(rsImages_var.Fields.Item("Expr1000").Value)%></font> -->
		  
				</div>
					
		</div>
		

</form>





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
