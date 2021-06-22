<%@language = "VBscript"%>
<% Response.Buffer=True %>
<%
	Dim conn
	set conn=server.CreateObject("ADODB.Connection")
	conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& Server.MapPath("store.mdb")
	Dim rs, strsql
	Set rs = Server.CreateObject("ADODB.Recordset")
	dim ID
	ID=Trim(Request.QueryString("ID"))
	If ID="" Then 
		Response.End
	End If
	ID=Cint(ID)
	strsql="Select * From Goods Where Gid Like '%"&ID&"%'"
	rs.Open strsql, conn, 1, 1
	Dim name
	name=Request.Cookies("user")("name")
    Response.Cookies("user")("name")=name
%>
<html>
	<head>
		<meta charset = "utf-8">
	</head>
	<body>
		<br>关于"<%=rs("GName")%>"的评论</h1>
		<nav>
			<a href="shopshow.asp" class="white">返回</a>
		</nav>
		<iframe name="message" src="input.asp?id=<%=ID%>" width="1400px" height="100px"></iframe>
		<iframe name="say" src="main.asp?id=<%=ID%>" width="1300px" height="400px"></iframe>
</body>
</html>