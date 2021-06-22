 <%Option Explicit%>
 <% Response.Buffer=True %>
 <%
    Dim name, email
    Response.Cookies("user")("ans")=""
    Response.Cookies("user")("verifyname")=""
    name=Request.Cookies("user")("verifyname")
    email=Request.Cookies("user")("verifyemail")
    
    name=Request("verifyname")
    email=Request("verifyemail")

    Response.Cookies("user")("verifyname")=name
    Response.Cookies("user")("verifyemail")=email
    Response.Cookies("user").Expires=DateAdd("d", 30, Date())
%>
<%
    Dim conn, rs
    set conn=server.CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& Server.MapPath("store.mdb")
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open "Select * From Users", conn, 1, 1
%>
 <head>
 <meta charset="UTF-8">
  <meta name="Generator" content="EditPlus®">
  <meta name="Author" content="">
  <meta name="Keywords" content="">
  <meta name="Description" content="">
  <title>forget password</title>
 </head>
 <form action="findpass.asp" method="post">
<%
    Dim i
    For i=1 To rs.PageSize
        if rs.Eof Then Exit For
        if name=rs("name") then 
            if email<>rs("email") then 
                Response.Write "用户名与邮箱不匹配"
                exit for
            else %>
                问题：<%=rs("tip-password")%><br>
                请输入答案:<input type="text" name="ans"><br>
                请输入用户名<input type="text" name="verifyname"><br>
                <input type="submit" value="submit">
        <%exit for
    end if
    end if
        rs.MoveNext
    next
%>
</form>