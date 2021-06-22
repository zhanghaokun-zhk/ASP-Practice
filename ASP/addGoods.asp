<head>
	<title>AddGoods</title>
<meta charset="UTF-8">
  <meta name="Generator" content="EditPlus®">
  <meta name="Author" content="">
  <meta name="Keywords" content="">
  <meta name="Description" content="">
</head>
<%
  Response.Cookies("user")("GName")=""
  Response.Cookies("user")("Price")=""
  Response.Cookies("user")("GDScript")=""
  Response.Cookies("user")("Rebate")=""
  Response.Cookies("user")("pic")=""
  Response.Cookies("user")("link")=""
%>
<body>
  <form action="adding.asp" method="POST">
  <p>输入商品详情<br>
  <p>名称:<input type="text" name="GName">
  <p>价格：<input type="text" name="Price">
  <p>商品类别：<input type=radio name=GDScript value="green" checked>绿茶
   <input type=radio name=GDScript value="red">红茶
   <input type=radio name=GDScript value="puer">普洱茶
   <input type=radio name=GDScript value="flower">花茶
   <input type=radio name=GDScript value="mix">混合
  <p>折扣：<input type="text" name="Rebate">
  <p>图片：<input type="text" name="pic"><br>
  <p>链接：<input type="text" name="link"><br>
  <input type="submit" value="提交">
</body>