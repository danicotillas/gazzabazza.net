<%
  Option Explicit
%>
<!--#INCLUDE FILE="_captcha.asp"-->
<%
  Dim checktext
  checktext = request.QueryString("text")
  if len(checktext) = 0 then checktext = session("checktext")

  response.ContentType = "image/gif"
  response.binarywrite textToGIF(checktext)
%>