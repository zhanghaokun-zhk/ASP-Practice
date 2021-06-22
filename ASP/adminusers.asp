<%Option Explicit%>
<head>
	<title>Goods</title>
<meta charset="UTF-8">
  <meta name="Generator" content="EditPlus®">
  <meta name="Author" content="">
  <meta name="Keywords" content="">
  <meta name="Description" content="">
</head>
<%
	Dim conn
	set conn=server.CreateObject("ADODB.Connection")
	conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& Server.MapPath("store.mdb")
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From Users", conn, 1, 3
%>
<center>
<div align="center">
<table width="53%" id="table1">
	<th>name</th><th>gender</th><th>email</th><th>address</th>
<%Dim i
  For i=1 To rs.PageSize
    if rs.Eof Then Exit For
%>
<tr>
 		<td width="200"><%=rs("name")%></td>
		<td><%=rs("gender")%></td>
		<td><%=rs("email")%></td>
		<td><%=rs("address")%></td>
		<form method="post" action="deleteusers.asp?id=<%=rs("ID")%>">
		<td><input type=submit value="Delete"></td></form>
	</tr>
<%
	rs.MoveNext
	next
%>
	<nav>
	<a href="admin.asp">Back</a>
	</nav>
<head>
	<title>Goods</title>
</head>