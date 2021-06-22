<%Option Explicit%>
<% Response.Buffer=True %>
<%
	Dim pass,repass,name,tip,ans,gender,email,admincode
	name=Request.Cookies("user")("name")
	pass=Request.Cookies("user")("paswrd")
	repass=Request.Cookies("user")("re-paswrd")
	tip=Request.Cookies("user")("tip-paswrd")
	ans=Request.Cookies("user")("ans-paswrd")
	gender=Request.Cookies("user")("gender")
	email=Request.Cookies("user")("email")
	admincode=Request.Cookies("user")("admincode")

	tip=Request("tip-paswrd")
	ans=Request("ans-paswrd")
	gender=Request("gender")
	name=Request("name")
	pass=Request("paswrd")
	repass=Request("re-paswrd")
	email=Request("email")
	admincode=Request("admincode")

	Response.Cookies("user")("email")=email
	Response.Cookies("user")("tip-paswrd")=tip
	Response.Cookies("user")("ans-paswrd")=ans
	Response.Cookies("user")("gender")=gender
	Response.Cookies("user")("name")=name
	Response.Cookies("user")("passwrd")=pass
	Response.Cookies("user")("re-paswrd")=repass
	Response.Cookies("user")("admincode")=admincode
	Response.Cookies("user").Expires=DateAdd("d", 30, Date())
%>
<%
	Dim conn,db,constr
	set conn=server.CreateObject("ADODB.Connection")
	db="store.mdb"
	constr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"") 
	conn.Open constr
	Dim rs, strsql, rss
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From Users", conn, 1, 3
	Set rss = Server.CreateObject("ADODB.Recordset")
	rss.Open "Select * From admin", conn, 1, 3
%>
 <head>
  <title>Check password</title>
 </head>
 <body>
 <%If pass<>repass then%>
	<form action="reg.asp" method="post">
	两次密码不一致
	<input type="submit" value="确定"></form>
<%Else
	Dim i,vis
	vis=1
	For i=1 To rs.PageSize
		If rs.Eof Then 
			Exit For
		End if
			If rs("name")=name Then
				vis=0
			End If
			rs.MoveNext
	Next
  	If vis<>0 Then
  		if admincode="admin123" then
  			rss.addnew
  			rss("name")=name
  			rss("password")=pass
  			rss.update
  			rss.close
  			set rss=Nothing
  		else
  			rs.addnew
			rs("name")=name
			rs("password")=pass
			rs("tip-password")=tip
			rs("answer")=ans
			rs("gender")=gender
			rs("email")=email
			rs("VIP")=0
			rs.update
			rs.close
			Set rs=Nothing
		end if
		conn.close()
		Response.redirect "index.asp"
	Else
  		Response.redirect "invalid.asp"
	End if
 End If%>
 </body>
</html>
