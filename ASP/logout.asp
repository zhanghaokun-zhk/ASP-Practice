<%@language="VBScript"%>
<%
	For Each cookie in Response.Cookies
		Response.Cookies(cookie).Expires = DateAdd("d",-1,now())
	Next
	Response.redirect "index.asp"
%>