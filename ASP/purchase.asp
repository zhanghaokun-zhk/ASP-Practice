<%Option Explicit%>
<head>
<meta charset="UTF-8">
  <meta name="Generator" content="EditPlus®">
  <meta name="Author" content="">
  <meta name="Keywords" content="">
  <meta name="Description" content="">
</head>
<%
    Dim name
    name=Request.Cookies("user")("name")
    Response.Cookies("user")("name")=name
    Response.Cookies("user").Expires=DateAdd("d", 30, Date())
%>
<%
	Dim conn
	set conn=server.CreateObject("ADODB.Connection")
	conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& Server.MapPath("store.mdb")
	Dim rss
	Set rss = Server.CreateObject("ADODB.Recordset")
	rss.Open "Select * From Users Where name like'"&name&"'", conn, 1, 3
%>
<title>purchase</title>
<div align="center">
<table width="53%" id="table1">
<th>address</th>
<%
	dim intPage, i
	If Not rss.Bof And Not rss.Eof Then
		For i=1 To rss.PageSize
    		if rss.Eof Then Exit For
%>
<tr>
                  <td><input type="radio" value="<%=rss("address")%>"  name="address"></td>
                  <td><%=rss("address")%></td>
				  <td><a href="updateaddress.asp">update</a></td>
                 </tr>
<tr>
<%
rss.MoveNext
next
end if%>
<form action="success.asp" method="post">
<input type=submit name="submit" value="submit">
</form>