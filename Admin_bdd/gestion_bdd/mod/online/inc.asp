<% ' MODULE ONLINE pour AxS - Cr�ation du compteur

modTABLE = "AxS_online"

' Cr�ation de la table
if request("act")="addonline" then
	Conn.execute("CREATE TABLE [" & modTABLE & "] (IP Text(15),start Integer)")
end if

' Recherche de la table
if TABLE <> modTABLE then
	Set ObjTable = Conn.OpenSchema(20)
	Do while not ObjTable.eof
		if ObjTable("table_type")="TABLE" AND ObjTable("table_name")=modTABLE then
			TABLE = modTABLE
			exit do
		end if
		ObjTable.moveNext
	Loop
	ObjTable.close
	if TABLE = modTABLE then
		urlk = replace("&base=" & baseRep & BASE & "&table=" & TABLE," ","+")
		urlk2 = "?lng=" & LNG & "&module=" & request("module") & urlk
		%><script>location.replace("<%=URL&urlk2%>")</script>
	<%end if
end if

' Affichage
%>
<tr><td align=center colspan=6 nowrap>
	<table cellpadding=5 border cellspacing=0><tr><td bgcolor=#FFFFFF width=300 nowrap align=center>
		<b>Compteur de visiteurs en ligne</b><br><br>
<%
if TABLE = modTABLE then%>
		Pour utiliser ce compteur sur votre site, il suffit de copier le code suivant dans chacune de vos pages :
		<br>
		<textarea wrap=virtual cols=50 rows=3>&lt;script language="JavaScript" src="http://<%=request.serverVariables("SERVER_NAME")&replace(request.serverVariables("SCRIPT_NAME"),"elt/admin_access.asp","mod/online.asp")&"?base="&BASE%>">&lt;/script></textarea>
		<br>Test du compteur >> <b><script language="JavaScript" src="../mod/online.asp?base=<%=BASE%>"></script></b>
<%else%>
		Ce module vous permet de cr�er automatiquement un compteur de visiteurs connect�s sur votre site.
		<br><br>En cliquant sur le lien ci-dessous, une nouvelle table va �tre cr��e dans votre base <b><%=BASE%></b> et le script � utiliser sera g�n�r�.<br>
		<br><A href="<%=URL & urlk2%>&act=addonline">Cr�er le compteur</A>
<%end if%>
	</td></tr></table>
</td></tr>
