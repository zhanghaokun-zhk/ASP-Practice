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
 <head>
 <meta charset="UTF-8">
  <meta name="Generator" content="EditPlus®">
  <meta name="Author" content="">
  <meta name="Keywords" content="">
  <meta name="Description" content="">
  <title>login</title>
 </head>
<body>
<form action="" method="post">
<p>用户名:<input type="text" name="name"></p>
  <p>密码:<input type="password" name="paswrd"></p>

<canvas id="canvas" width="100" height="43" onclick="dj()" 
    style="border: 1px solid #ccc;
        border-radius: 5px;"></canvas><br>
          <input type="text" value="" placeholder="请输入验证码（区分大小写）" 
   style="height:23px; top:-15px; font-size:20px;" id ="text" name="text"><br>
    <button class="btn" onclick="sublim()">登录</button><br>
   <!--#include file="validcode2.asp"-->
   <%if Trim(request("name"))<>""  and Trim(request("text"))="lhywmzyqb" then
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
   		end if
    %>
 </form>
        <form action="forgetpass.asp" method="post">
    <input type=submit name="submit" value="忘记密码"></form>
</body>