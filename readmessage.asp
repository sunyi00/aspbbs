<!--#include file=conn.asp-->
<%
response.buffer=true
response.clear
response.expires=0
%>
<html>
<head>
</head>

<body bgcolor=#c1f7d8>
<p align=center><font size=5>Read Messages</font></p>
<%
if not request.querystring("id")="" then
  strdelesql="delete from message where id=" & request.querystring("id")
  set rsdelete=server.createobject("adodb.recordset")
  rsdelete.open strdeletesql,conn,1,3
  set rsdelete=nothing
  response.redirect "readmessage.asp"
end if
%>
<%
dim strsql
strsql="select * from message where messagetoname='" & session("name") & "'"
set rs=server.createobject("adodb.recordset")
rs.open strsql,conn,1,3
response.write "<center> you have " & rs.recordcount & "message(s)"
%>
<%
dim i
response.write "<table border=1 width=100%/><tr>"
for i=1 to rs.recordcount
  response.write "<td width=27%/>" & rs("messagedate") & "    " & rs("messagetime") & "</td>"
  response.write "<td width=18%/>" & rs("messagename") & "</td>"
  response.write "<td width=20%/>" & rs("messageemail") & "</td>"
  response.write "<td width=20%/>" & rs("messagehomepage") & "</td>"

response.write "<td width=15%/><a href=return.asp?toname=" & rs("messagename")
response.write ">reply</a>" & "   "
response.write "<a href=readmessage.asp?id=" & rs("id") & ">delete</a></td></tr>"
response.write "<tr><td colspan=5>" & rs("messagecontent") & "</td></tr>"
rs.movenext
next
response.write "</table>"
rs.close
set rs=nothing
%>
</body>
</html>
