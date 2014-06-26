<%
'ASP Captcha include.
' 2006 Antonin Foller, Motobit Software
' http://www.motobit.com/
' This code is freeware.
'Please include a link to www.motobit.com page 
' on your pages if you use this include.

Sub CreateGifFromText(inText, FileName)

  'Create an OWC object
  Dim chs
  Set Chs = getOWC
  

  'Get chart constants
  Dim chConstants: Set chConstants = chs.Constants
  
  'Get a chart object 
  Dim Chart: Set Chart = chs.Charts.Add

  'Enable title for the chart.
  Chart.HasTitle = True
	
	randomize

  'Set the text and properties.  
  Chart.Title.Caption = inText

  'set random fonts.
	Dim Fonts, FontSizeAdd
	FontSizeAdd = int(rnd * 10)
	Fonts = array("Times New Roman","Arial","Book Antiqua","Comic Sans MS","Haettenschweiler","Lucida Console","Monotype Corsiva","Impact")
  Chart.Title.Font.Name = Fonts(rnd * ubound(Fonts))
  Chart.Title.Font.Size = FontSizeAdd + 13
	Chart.Title.Font.Color = rnd * &H1000000
	if rnd>0.5 then Chart.Title.Font.italic = true
	if rnd>0.5 then Chart.Title.Font.bold = true
  
  'Set some chart background 
  '(Interior of the ChartSpace and Title)
	do 
	  on error resume next
    chs.Interior.SetPresetGradient int(1 + rnd * 7), _
      int(1 + rnd * 4), int(1 + rnd * 24)
		Chart.Title.Interior.SetPresetGradient int(1 + rnd * 7), _
      int(1 + rnd * 4), int(1 + rnd * 24)
  loop while err<>0
	on error goto 0

  'Save the image as a file
  chs.ExportPicture FileName, , 10 + 20*len(intext) + 4 * FontSizeAdd , 45 + 1.5 * FontSizeAdd 
End Sub

Function getOWC
  On error resume next
  Dim chs
  Set Chs = CreateObject("OWC10.ChartSpace") ' As New ChartSpace
  if isempty(Chs) then Set Chs = CreateObject("OWC11.ChartSpace") 
  'if isempty(Chs) then Set Chs = CreateObject("OWC.Chart") 
  Set getOWC = Chs
End Function 

'http://www.motobit.com/tips/detpg_read-write-binary-files/
Function ReadBinaryFile(FileName)
  Const adTypeBinary = 1
  
  'Create Stream object
  Dim BinaryStream
  Set BinaryStream = CreateObject("ADODB.Stream")
  
  'Specify stream type - we want To get binary data.
  BinaryStream.Type = adTypeBinary
  
  'Open the stream
  BinaryStream.Open
  
  'Load the file data from disk To stream object
  BinaryStream.LoadFromFile FileName
  
  'Open the stream And get binary data from the object
  ReadBinaryFile = BinaryStream.Read
End Function


Function GetTempFileName(Byref FS)
  randomize
  GetTempFileName = FS.GetSpecialFolder(2)  & "\" & rnd & ".gif"
End Function 

Function RandomText(Length)	
	Dim I, Out
	Randomize
	For I = 1 to Length
		Out = Out & Chr(64 + rnd * 28) 
	Next
	RandomText = Out 
End Function

Function textToGIF(inText)
	Dim FS: Set FS = CreateObject("Scripting.FileSystemObject")

  'get a temporary file name
  Dim FileName: FileName = GetTempFileName(FS)

  'Create the GIF file with a text.
  CreateGifFromText inText, FileName

  'Get the file as a binary data from disk
  textToGIF = ReadBinaryFile(FileName)

  'Delete the temporary file
  FS.DeleteFile FileName
End Function


%>