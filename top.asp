<%@language=vbscript%>
<html>
<head>
<meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
</head>

<body bgcolor=#c1f7d8>
<br>
<base target=main>
<table border=1 width=100%>
  <tr>
    <td width=15% align=center>
      <%response.write session("name")%>
    </td>
    <td width=17% align=center>
      <a href=changepassword.htm>change password</a>
    </td>
    <td width=17% align=center>
      <a href=readmessage.asp>read message</a>
    </td>
    <td width=17% align=center>
      <a href=showlist.asp>article list</a>
    </td>
    <td width=17% align=center>
      <a href=publisharticle.htm>publish article</a>
    </td>
    <td width=17% align=center>
      <a href="/" target=_parent>back to home</a>
    </td>
  </tr>
  <tr>
    <td colspan=6 align=center>welcome</td>
  </tr>
</table>
</body>
</html>
