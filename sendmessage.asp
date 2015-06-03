<%@language=vbscript%>
<!--#include file=conn.asp-->
<html>
<head>
</head>

<body bgcolor=#c1f7d8>
<center>
<%
dim strname,strcontent,strtable
strname=request.form("name")
strcontent=request.form("content")
strtoname=request.form("toname")
if trim(strname)="" or trim("strcontent")="" then
  response.write "<br><br>name and content should not be null"
%>
<br><br><a href=javascript.history.back()>back</a><br>
<%
  response.end
end if
strtable="message"
set rs=server.createobject("adodb.recordset")
rs.open strtable,conn,1,3
rs.addnew
rs("messagename")=strname
rs("messagecontent")=strcontent
rs("messagetoname")=strtoname
rs.update
rs.close
set rs=nothing
response.write "con, your message was sent to " & strtoname
%>
</center>
</body>
</html>
