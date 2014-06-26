<% ' ±AxS : Export MSAccess >> Text

temp = ""
for each Champ in RS.fields
	temp = temp & Champ.name & sep1
next
temp = temp & sep2
inF.writeline(temp)
while not RS.eof
	temp = ""
	for each Champ in RS.fields
		if not isnull(Champ.value) then
			select case Champ.type
				case 2,3,4,5,6,17 : temp = temp & replace(cStr(Champ.value),",",".")
				case 11 : temp = temp & cStr(Champ.value)
				case 135,200,201,202,203 :
					temp2 = replace(Champ.value,sep1,sep1b)
					temp = temp & replace(temp2,sep2,sep2b)
				case else :
			end select
		end if
		temp = temp & sep1
	next
	temp = temp & sep2
	inF.writeline(temp)
	RS.movenext
wend
%>