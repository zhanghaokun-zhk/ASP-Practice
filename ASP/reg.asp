 <%
	Response.Cookies("user")("paswrd")=""
	Response.Cookies("user")("re-paswrd")=""
	Response.Cookies("user")("name")=""
 %>
<!DOCTYPE html>
<html lang="en">
 <head>
  <meta charset="UTF-8">
  <meta name="Generator" content="EditPlus®">
  <meta name="Author" content="">
  <meta name="Keywords" content="">
  <meta name="Description" content="">
  <title>Document</title>
 </head>
 <body>
 <form name="reg.asp" action="checkpass.asp" method="POST">
  <p>会员注册<br>
  <p>用户名:<input type="text" name="Name">
  <p>密码：<input type="password" name="paswrd">
  <p>密码确认：<input type="password" name="re-paswrd">
  <p>密码提示：<input type="text" name="tip-paswrd">
  <p>密码回答：<input type="text" name="ans-paswrd"><br>
  <p>性别：<input type=radio name=gender value="male" checked>男
   <input type=radio name=gender value="female">女
  <p>邮箱：<input type="text" name="email"><br>
    <p>admin code:<input type="admincode" name="admincode"><br>
  <input type="submit" value="提交">
 </body>
</html>
