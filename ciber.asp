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


<meta http-equiv="Page-Exit" content="blendTrans(Duration=1)">
<META HTTP-EQUIV="Refresh" CONTENT="10">


<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="description" content="La raz&oacute;n nos limita, el azar es INFINITO. Es &eacute;l quien nos descubrir&aacute; los sentidos ocultos, los sabores desconocidos, los placeres an&oacute;nimos que esconden las im&aacute;genes cotidianas. Sus significados son m&uacute;ltiples y extra&ntilde;os, y es aqu&iacute;, cuando definitivamente, encontrar&aacute;s la llave." />
<meta name="verify-v1" content="x8SGvOjKSa0+DPztRByBbbPpVCD4J7mANcCQIn/48SY=" />

<title>Gazza &amp; Bazza ::: El Azar Manda en Nuestras Vidas</title>
<link href="css_randomness/gazzabazza3.css" rel="stylesheet" type="text/css">


</head>

<body>


	
		<div class="container">
		
<div class="profoto">

<img src="img/<%=(rImagen.Fields.Item("imagen").Value)%>" align="middle" style=" min-width:300px; max-width:950px; max-height:650px;" />
							
              
								<div class="texto">
								  
								      <%response.Write((rTexto.Fields.Item("ideas").Value))%>
						   
								</div>
					
              
								
            
            		
				</div>
         

                  
				  
				<div class="recontainer">
							
						<div class="boton" align="center">
		 							
								<div class="contact" align="center">
           							
								  <div align="center">
										
           								<p><span class="linkcontact"><img src="images_randomness/gb_ciberin.png" width="27" height="23" border="0" align="absmiddle" class="ostiaputa"></span>            							  <br>
       								           								      <font color="#FFFFFF" size="2"></span>
       								                                <strong>www.gazzabazza.net</strong><span class="barre"></font><br>
   	                </p>
   								  </div>
										
								</div>
									
						</div>
	  
	                  
		  
				</div>
					
		</div>
		

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
