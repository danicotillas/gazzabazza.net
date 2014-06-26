<!--#include file="chaine.asp"-->
<!--#include file="o_ami_upload.asp" --->
<HTML><HEAD>
	<TITLE>±AxS - Upload</TITLE>
</HEAD>
<BODY BGCOLOR="white" onunload="opener.location.reload()">
<%
set upload = new ami_upLoad
'----- paramètres pour l'upload (facultatif)
upload.extensionsUploadee("txt")
'upload.extensionsNonUploadee("html txt doc htm")
upload.tailleFichiersUploades("3000")
'upload.repertoireServeur("\test_dechargt\")

a=upload.upload(Request.BinaryRead(Request.TotalBytes),Request.ServerVariables("HTTP_Content_Type"))
	' trace(upload.NbreFichiersEcrits()&" fichiers écrits sur "&upload.NbreTotalFichiers())

	' upload.afficheObjet()

	'liste des fichiers uploadés
	'trace("")
	'trace("liste des fichiers uploadés :")
	'trace("-----------------------------")
	'set courant = upload.fichier
	'while not courant.estNull
	'	if courant.estEcrit() then trace(courant.repertoire&courant.nom)
	'	set courant=courant.suivant
	'wend 

	'liste des fichiers non uploadés
	'trace("")
	'trace("liste des fichiers non uploadés :")
	'trace("---------------------------------")
	'tab = upload.fichiersNonEcrits()
	'i=0
	'while tab(i) <> ""
	'	trace(tab(i))
	'	i=i+1
	'wend 
	
'else
'	trace("aucun fichier transmis pour upload!!!")
'end if
set upload = nothing
%>
</BODY>
</HTML>
<script>window.close()</script>