<% ' ±AxS : Export MSAccess >> Excel (.slk)

function txt2slk(txt)
	dim temp
	if not isnull(txt) then
		temp = replace(txt,VbCrLf," ")
		txt2slk = replace(temp,chr(13)," ")
	end if
end function

temp = "ID;PWXL;N;E" & VbCrLf _
	& "P;PGeneral" & VbCrLf
	select case lng
		case "fr" : temp = temp & "P;Pdd/mm/yyyy\ hh:mm:ss" & VbCrLf
		case else : temp = temp & "P;Pmm/dd/yyyy\ hh:mm:ss" & VbCrLf
	end select
Y = 1 : X = 1 : temp = temp & "C;Y" & Y & ";"
for each Champ in RS.fields
	if X > 1 then temp = temp & "C;"
	temp = temp & "X" & X & ";K" & chr(34) & Champ.name & chr(34) & VbCrLf
	X = X+1
next
inF.writeLine(temp)
while not RS.eof
	temp = ""
	Y = Y+1 : X = 1 : temp = temp & "C;Y" & Y & ";"
	for each Champ in RS.fields
		if Champ.type = 135 then temp = temp & "F;P1;FG0G;X" & X & VbCrLf
		if X > 1 then temp = temp & "C;"
		if Champ.type <> 135 then temp = temp &  "X" & X & ";"
		temp = temp & "K"
		if not isnull(Champ.value) then
			select case Champ.type
				case 2,3,4,5,6,17 : temp = temp & replace(cStr(Champ.value),",",".")
				case 11 : temp = temp & chr(34) & cStr(Champ.value) & chr(34)
				case 135 :
temp = temp & replace(cStr(datediff("s","30/12/1899",Champ.value)/(3600*24)),",",".")
				case 200,202 : temp = temp & chr(34) & txt2slk(Champ.value) & chr(34)
				case 201,203 : temp = temp & chr(34) & left(txt2slk(Champ.value),255) & chr(34)
				case else : temp = temp & chr(34) & chr(34)
			end select
		else
			temp = temp & chr(34) & chr(34)
		end if
		temp = temp & VbCrLf
		X = X+1
	next
	inF.writeLine(temp)
	RS.movenext
wend
inF.writeLine("E" & VbCrLf)
%>