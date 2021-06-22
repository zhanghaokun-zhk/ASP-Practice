  <% Response.Buffer=True %>
  <%
    Response.Cookies("user")("verifyemail")=""
    Response.Cookies("user")("verifyname")=""
 %>
 <head>
 <meta charset="UTF-8">
  <meta name="Generator" content="EditPlus®">
  <meta name="Author" content="">
  <meta name="Keywords" content="">
  <meta name="Description" content="">
  <title>forget password</title>
 </head>
<body>
    <form action="verifyhint.asp" method="post">
        请输入用户名<input type="text" name="verifyname"><br>
        请输入邮箱<input type="text" name="verifyemail"><br>
        <input type="submit" value="submit">
    </form>
</body>