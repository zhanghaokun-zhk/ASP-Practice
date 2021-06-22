 <%@language="VBSCRIPT"%>
 <%Option Explicit%>
<form action="shopshow.asp" method="post">
<% Response.Buffer=True %>
<%
	Dim name,gender,priceup,pricedown
	name=Request.Cookies("user")("name")
	gender=Request.Cookies("user")("gender")
	priceup=Request.Cookies("user")("priceup")
	Response.Cookies("user")("name")=name
	Response.Cookies("user")("gender")=gender
	Response.Cookies("user")("priceup")=priceup
	Response.Cookies("user").Expires=DateAdd("d", 30, Date())
%>
<html>
<head>
<meta charset="UTF-8">
  <meta name="Generator" content="EditPlus®">
  <meta name="Author" content="">
  <meta name="Keywords" content="">
  <meta name="Description" content="">
</head>
<title>show</title>
 <body>
 <h2 align="center">查找</h2>
 <%="欢迎" &name %>
 <% If gender="male" then%>
 先生光临<br>
 <%Else%>
 女士光临<br>
 <%End If%>
 <%
	Dim conn, rs, rsss, strsql
	set conn=server.CreateObject("ADODB.Connection")
	conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& Server.MapPath("store.mdb")
	Set rs = Server.CreateObject("ADODB.Recordset")
	set rsss = Server.CreateObject("ADODB.Recordset")
	rsss.Open "Select * From Users", conn, 1, 1
	if priceup = "1" then
		strsql = "Select * From Goods Order By Price"
	else if priceup = "2" then
					strsql = "Select * From Goods Order By Price DESC"
	 		 else 
	 		 strsql="Select * from Goods"
	 end if
	end if
	If Request.Form("txtName")<>""Then
	  strsql = "Select * From Goods Where GName Like '%"&Request.Form("txtName")&"%'"
	 end if
	 if Request.form("class") <> "null" then
	 	strsql="select * from Goods Where GDScript like '%"&Request.Form("class")&"%'"
	 end if
	 rs.Open strsql, conn, 1, 1
%>
 <form name="frmSearch" method="POST" action="">
   请输入要查找的物品：<input type="text" name="txtName">
   <select name="class">
   	<option value="null">请选择</option>
   	<option value="green">绿茶</option>
   	<option value="red">红茶</option>
   	<option value="flower">花茶</option>
   	<option value="mix">混合</option>
   	<option value="puer">普洱茶</option>
   </select>
      <input type="submit" name="btnSubmit" value="确定">
  </form>
 <table border="1" width="100%" align="center">
 <tr bgcolor="#E0E0E0">
<nav>
<a href="priceup.asp">价格升序</a>
<a href="pricedown.asp">价格降序</a>
<a href="cart.asp">购物车</a>
<a href="address.asp">收货地址</a>
<a href="aboutus.asp">关于我们</a>
<a href="logout.asp">退出登录</a>
</nav>
<%
  If Not rs.Bof And Not rs.Eof Then
	Dim intpage
	If Request.QueryString("varPage")="" Then
		Intpage=1
	Else
	  intPage=CInt(Request.QueryString("varPage"))
	End If
	rs.PageSize=5
	rs.AbsolutePage=intPage
%>
<%Dim goodsnum(100)%>
<div align="center">
<table width="53%" id="table1">
<th>Gid</th><th>pic</th><th>GName</th><th>Price</th><th>num</th>
<%Dim i, flag,j
flag=0
  For i=1 To rs.PageSize
    if rs.Eof Then Exit For
%>
 <tr>
 		<td><%=rs("Gid")%></td>
 		<td><a href=<%=rs("link")%>><img src="<%=rs("pic")%>" width="200px" height="200px"></td> 
 		<td width="200"><a href="comment.asp?id=<%=rs("Gid")%>"><%=rs("GName")%></td></a>
		<td><%for j=1 to rsss.PageSize
		        if rsss.Eof then exit for
		if rsss("name")=name and cint(rsss("VIP"))=1 then flag=1
			rsss.MoveNext
		next
			if flag=1 then%>
				原价￥<%=rs("Price")%><br>折后￥<%=rs("Price")*(rs("Rebate"))%>
					<%else%>
					￥<%=rs("Price")%>
					<%end if%>
				</td>
		<form method="post" action="insert.asp?id=<%=rs("Gid")%>">
		<td width="70"><input type=text name="Text<%=rs("Gid")%>" size="8"></td>
		<td width="70"><input type=submit name="submit" value="add to cart"></td>
		<%Dim insertasp
			insertasp="insert.asp?id='"&i&"'"%>
	</tr>
	<%
		Dim rss,Num
		Num=Request.Cookies("user")("num")
		Response.Cookies("user")("num")=Num
		Response.Cookies("user").Expires=DateAdd("d", 30, Date())
		Set rss = Server.CreateObject("ADODB.Recordset")
		rss.Open "Select * From Cart", conn, 1, 3
		If Num<>"" Then
			rss.addnew
			rss("GName")=rs("GName")
			rss("Price")=rs("Price")
			rss("Num")=CInt(Num)
		End If%></form><%
		rs.MoveNext
		Next
		End If
	%>
	<%
	  If intpage>1 Then
		Response.Write "<a href='shopshow.asp?varPage=" &(intPage-1)&"'>上一页</a>"
	  Else
	    Response.Write "上一页&nbsp;"
	  End If
	  If intPage < rs.PageCount Then
		Response.Write "<a href='shopshow.asp?varPage="&(intPage+1)&"'>下一页</a>"
	  Else
        Response.Write "下一页&nbsp;"
	  End If
	%>
 </body>
 </html>