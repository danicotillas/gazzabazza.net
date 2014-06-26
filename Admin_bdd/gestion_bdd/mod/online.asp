<% ' MODULE ONLINE pour AxS - Appel du compteur

TABLE = "AxS_online"
BASE = request("base")
' IP du visiteur
IP=request.serverVariables("REMOTE_ADDR")
' Date/heure courante en minutes
Session.LCID = 1036
date0 = dateDiff("n","05/01/2003",now)
' Dur�e de vie max
vie = 5

%><!-- #include file="../_params.asp"--><%

if BASE <> "" then
	basePath = replace(user_baseFolder & BASE & ".mdb","/","\")
	' Connexion � la BD
	Set Conn = Server.CreateObject("ADODB.Connection")
	Set RS = Server.CreateObject("ADODB.Recordset")
	Conn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & basePath

	' Suppression des anciens
	Conn.execute "DELETE * FROM " & TABLE & " WHERE start<" & (date0-vie)

	' Stockage du hit courant
	SQL = "SELECT * FROM " & TABLE & " WHERE IP='" & IP & "'"
	RS.open SQL,Conn,3,3
	' Si non stock�, on ajoute
	if RS.eof then RS.addnew
	' Mise � jour
	RS("IP") = IP
	RS("start") = date0
	RS.update
	RS.close

	' Nombre de visiteurs en ligne
	SQL = "SELECT count(IP) FROM " & TABLE
	RS.open SQL,Conn,3,3
	online = RS(0)
	RS.close

	' D�connnexion
	Conn.close

	' Retour javascript
	JVS = online & " en ligne"
end if
%>document.write("<%=JVS%>");