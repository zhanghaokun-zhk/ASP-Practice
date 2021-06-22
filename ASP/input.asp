<%@language = "VBscript" codepage = "65001"%>
<% Response.Buffer=True %>
<%Response.charset="utf-8"%>
<html>
<head runat="server">
</head>
<body>
<%
    Dim conn
    set conn=server.CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& Server.MapPath("store.mdb")
    Dim rs
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open "Select * From Comment", conn, 1, 3
%>
<%
	Response.Write Guest
	dim ID
	ID=Trim(Request.QueryString("ID"))
	If ID="" Then Response.End
	ID=Cint(ID)
	Dim Guest
	Guest=Request.Cookies("user")("name")
    Response.Cookies("user")("name")=Guest
	%>
	<form name="form" method="post" action="">
		<p>发言:<input type="text" name="Say" size=50></p>
		<p><input type="submit" value="发送"></p>
	</form>
<%
	Response.Write saying
	saying = Request.Form("Say")
	if saying <>  "" then
		rs.addnew
		rs("Guest")=Guest
		rs("Date")=Date()
		rs("Gid")=ID
		rs("Comment")=saying
		rs.update
		rs.close
	end if
%>
</body>
</html>