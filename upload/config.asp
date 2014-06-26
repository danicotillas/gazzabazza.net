
	<%	
	'-----------------------------------------------
	'name of folder (note there is no / at end)
	strFolder = Server.MapPath("/img")
	'name of folder in http format (note there is no / at end)
	httpRef = ("img")
	'the max size of file which can be uploaded, 0 will give unlimited file size

	lngFileSize = 250000

	'the files to be excluded (must be in format ".aaa;.bbb")
	'and must be set to blank ("") if none are to be excluded
	strExcludes = ""
	'the files to be included (must be in format ".aaa;.bbb")
	'and must be set to blank ("") if none are to be excluded
	strIncludes = ".gif;.jpg;.bmp;.jpeg;.GIF;.JPG;.BMP;.JPEG"
	'-----------------------------------------------
	
	%>
	