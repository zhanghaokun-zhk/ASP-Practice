<%@language="VBscript" codepage="65001"%>
<%Response.charset="utf-8"%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
	<title></title>
	<meta charset="utf-8">
	<meta http-equiv="refresh" content="2">
</head>
<%
    Dim conn
    set conn=server.CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& Server.MapPath("store.mdb")
    Dim rs
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open "Select * From Comment", conn, 1, 3
    dim ID
	ID=Trim(Request.QueryString("ID"))
%>
<body>
	<%
		dim saying
		rs.PageSize = 1000
		for i=1 to rs.PageSize
			if rs.Eof then exit For
			if rs("Gid")=ID then saying=rs("Date")&"-"&rs("Guest")&"说:"&rs("Comment")&"<br>"
			Response.Write saying
			rs.MoveNext
		next
	%>	
</body>
</html>