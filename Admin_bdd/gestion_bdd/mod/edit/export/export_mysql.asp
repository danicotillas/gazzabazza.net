<% ' ±AxS : Export MSAccess >> MySQL

function txt2mys(txt)
	dim temp
	if not isnull(txt) then
		temp = replace(txt,VbCrLf,"\r\n")
		temp = replace(temp,chr(13),"\r\n")
		txt2mys = replace(temp,"'","\'")
	end if
end function

function n2t(n,l)
	n2t = right("000" & cStr(n),l)
end function

temp = "#" & VbCrLf _
	& "# Export Access >> MySQL" & VbCrLf _
	& "# with " & Version & VbCrLf _
	& "#" & VbCrLf _
	& "# DB : " & BASE & VbCrLf _
	& "# SQL : " & SQL & VbCrLf _
	& "#" & VbCrLf& VbCrLf
 
temp = temp & "DROP TABLE IF EXISTS `" & TABLE & "`;" & VbCrLf _
	& "CREATE TABLE `" & TABLE & "` (" & VbCrLf
for each Champ in RS.fields
	temp2 = "  `" & Champ.name & "` "
	select case Champ.type
		case 2 : temp = temp & temp2 & "smallint," & VbCrLf
		case 3 : temp = temp & temp2 & "int," & VbCrLf
		case 4 : temp = temp & temp2 & "float," & VbCrLf
		case 5 : temp = temp & temp2 & "double," & VbCrLf
		case 6 : temp = temp & temp2 & "float(15,4)," & VbCrLf
		case 11 : temp = temp & temp2 & "char(1)," & VbCrLf
		case 17 : temp = temp & temp2 & "tinyint," & VbCrLf
		case 135 : temp = temp & temp2 & "datetime," & VbCrLf
		case 200,202 : temp = temp & temp2 & "varchar(" & Champ.definedsize & ")," & VbCrLf
		case 201,203 : temp = temp & temp2 & "longtext," & VbCrLf
	end select
next
temp = left(temp,len(temp)-3) & VbCrLf & ") TYPE=MyISAM;" & VbCrLf & VbCrLf
inF.writeline(temp)
while not RS.eof
	temp = "INSERT INTO `" & TABLE & "` ("
	for each Champ in RS.fields
		select case Champ.type
			case 2,3,4,5,6,11,17,135,200,201,202,203 : temp = temp & "`" & Champ.name & "`,"
		end select
	next
	temp = left(temp,len(temp)-1) & ") VALUES ("
	for each Champ in RS.fields
		select case Champ.type
			case 2,3,4,5,6,17 : temp = temp & Champ.value & ","
			case 11 : 
				if(Champ.value) then temp = temp & "'Y'," _
				else temp = temp & "'N',"
			case 135 : temp = temp & "'" _
				& n2t(year(Champ.value),4) & "-" _
				& n2t(month(Champ.value),2) & "-" _
				& n2t(day(Champ.value),2) & " " _
				& n2t(hour(Champ.value),2) & ":" _
				& n2t(minute(Champ.value),2) & ":" _
				& n2t(second(Champ.value),2) & "',"
			case 200,201,202,203 : temp = temp & "'" & txt2mys(Champ.value) & "',"
		end select
	next
	temp = left(temp,len(temp)-1) & ");" & VbCrLf
	inF.writeline(temp)
	RS.movenext
wend
%>