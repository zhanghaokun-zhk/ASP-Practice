<%@language = "VBscript"%>
<%Option Explicit%>
<% Response.Buffer=True %>
<%
	Dim pass,repass,name,tip,ans,gender,email
	name=Request.Cookies("user")("name")
	pass=Request.Cookies("user")("paswrd")

	name=Request("name")
	pass=Request("paswrd")

	Response.Cookies("user")("name")=name
	Response.Cookies("user")("passwrd")=pass
	Response.Cookies("user").Expires=DateAdd("d", 30, Date())
%>
<%
	Dim conn,db,constr
	set conn=server.CreateObject("ADODB.Connection")
	db="store.mdb"
	constr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"") 
	conn.Open constr
	Dim rs, strsql, rss, vis, flag, i
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From Users", conn, 1, 3
	Set rss = Server.CreateObject("ADODB.Recordset")
	rss.Open "Select * from admin", conn, 1, 3
	vis=1
	flag=1
	For i=1 To rs.PageSize
		If rs.Eof Then 
			Exit For
		End if
			If rs("name")=name And rs("password")=pass Then
				vis=0
			End If
			rs.MoveNext
	Next
	If vis=1 Then
		for i=1 to rss.PageSize
			if rss.Eof then
				Exit For
			End If
			if rss("name")=name and rss("password")=pass then
				flag=0
			end if
			rss.MoveNext
		next
		if flag=0 then 
			response.Redirect "admin.asp"
		else
			Response.redirect "error.asp"
		end if
	Else
		Response.redirect "shopshow.asp"
	End If
%>
<body>
</body>