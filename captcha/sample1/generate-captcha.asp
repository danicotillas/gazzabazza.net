<%
  Option Explicit
%>
<!--#INCLUDE FILE="_captcha.asp"-->
<%
  response.ContentType = "image/gif"
  response.binarywrite textToGIF(session("checktext"))
%>