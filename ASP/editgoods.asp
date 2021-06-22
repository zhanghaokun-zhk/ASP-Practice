<%Option Explicit%>
 <head>
 <meta charset="UTF-8">
  <meta name="Generator" content="EditPlus®">
  <meta name="Author" content="">
  <meta name="Keywords" content="">
  <meta name="Description" content="">
 </head>
<%
	dim i, ID, strsql
	ID=cint(Trim(Request.QueryString("ID")))
	Dim conn
	set conn=server.CreateObject("ADODB.Connection")
	conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& Server.MapPath("store.mdb")
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "SELECT * FROM Goods", conn, 1, 3
%>
<div align="center">
<table width="53%" id="table1">
<th>Gid</th><th>pic</th><th>GName</th><th>Price</th><th>Rebate</th><th>New Price</th><th>New Rebate</th>
<%
	strsql="Select * From Goods Where Gid = '"&cint(ID)&"'"
	set rs = conn.Execute(strsql)
	for i=1 to rs.PageSize
		if rs.Eof then exit for
%>
<tr>
 		<td><%=rs("Gid")%></td>
 		<td><a href=<%=rs("link")%>><img src="<%=rs("pic")%>" width="200px" height="200px"></td> 
 		<td width="200"><%=rs("GName")%></td>
		<td><%=rs("Price")%></td>
		<td><%=rs("Rebate")%></td>
		<form method="post" action="price.asp?id=<%=rs("Gid")%>">
		<td width="70"><input type=text name="Price<%=rs("Gid")%>" size="8"></td>
		<td width="70"><input type=text name="Rebate<%=rs("Gid")%>" size="8"></td>
		<td width="70"><input type=submit name="submit" value="submit"></td>
</tr>
<%
	rs.MoveNext	
	Next
%>