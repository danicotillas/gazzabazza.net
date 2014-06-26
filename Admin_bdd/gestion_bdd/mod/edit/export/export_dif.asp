<% ' ±AxS : Export MSAccess >> Excel (.tif)

Function RC(strtxt)
   RC = replace (strtxt, "â", "ƒ")
   RC = replace (RC, "ä", "„")
   RC = replace (RC, "à", "…")
   RC = replace (RC, "é", "‚")
   RC = replace (RC, "è", "Š")
   RC = replace (RC, "ê", "ˆ")
   RC = replace (RC, "ë", "‰")
   RC = replace (RC, "ï", "‹")
   RC = replace (RC, "î", "Œ")
   RC = replace (RC, "ô", "o")
   RC = replace (RC, "ö", "”")
   RC = replace (RC, "ü", "")
   RC = replace (RC, "ù", "—")
   RC = replace (RC, "û", "–")
   RC = replace (RC, "ç", "‡")
   RC = replace (RC, "ñ", "¤")
   RC = replace (RC, VbCrLf, " ")
End Function

inF.writeLine "TABLE"
inF.writeLine "0,1"
inF.writeLine chr(34)&"EXCEL"&chr(34)
inF.writeLine "VECTORS"
inF.writeLine "0," & RS.recordcount ' nb de lignes
inF.writeLine chr(34)&chr(34)
inF.writeLine "TUPLES"
inF.writeLine "0," & RS.fields.count ' nb de colonnes
inF.writeLine chr(34)&chr(34)
inF.writeLine "DATA"
inF.writeLine "0,0"
inF.writeLine chr(34)&chr(34)
inF.writeLine "-1,0"
inF.writeLine "BOT"
for each Champ in RS.fields
	inF.writeLine "1,0"
	inF.writeLine chr(34) & Champ.name & chr(34)
next
inF.writeLine "-1,0"
while not RS.eof
	inF.writeLine "BOT"
	for each Champ in RS.fields
		inF.writeLine "1,0"
		select case Champ.type
			case 200,201,202,203 : 
				inF.writeLine chr(34) & RC(RS(Champ.name)) & chr(34)
			case else:
				inF.writeLine chr(34) & RS(Champ.name) & chr(34)
		end select
    next
    inF.writeLine "-1,0"
	RS.movenext
wend
inF.writeLine "EOD"
%>