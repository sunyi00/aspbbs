<%@language=vbscript%>
<!--#include file=conn.asp-->

<html>
<head>
<meta name="generator" content="Microsoft Visual Studio 6.0">
</head>
<body bgcolor=#c1f7d8>
<p align=center>
  <font size=5>Read Article</font>
</p>

<%
dim strselectsql
set rs=server.createobject("adodb.recordset")
strselectsql="select * from article where articleid=" & request.querystring("id")
rs.open strselectsql,conn,3,1
%>

<table border=1 width=100%>
  <tr>
    <td align=center width=25%>Date</td>
    <td align=center width=49%>Title</td>
    <td align=center width=10%>Author</td>
    <td align=center width=8%>Read</td>
    <td align=center width=8%>Fellow</td>
  </tr>
  <tr>
    <td align=center><%=rs("articledate")%>&nbsp;<%=rs("articletime")%></td>
    <td align=center><%=rs("articletitle")%></td>
    <td align=center><%=rs("articleauthor")%></td>
    <td align=center><%=rs("articleaccessnumber")%></td>
    <td align=center><%=rs("articlefellownumber")%></td>
  </tr>
</table>

<br>
Article Content
<br><br>
<%=rs("articlecontent")%>

<%
rs.close
set rs=nothing
%>

<%
strchangesql="update article set articleaccessnumber=articleaccessnumber+1 where articleid=" & Request.querystring("id")
set changers=server.createobject("adodb.recordset")
changers.open strchangesql,conn,1,3
set changers=nothing
%>

<%
strselectsql="select * from article where articleparent=" & request.querystring("id")
set rs=server.createobject("adodb.recordset")
rs.open strselectsql,conn,3,1
rs.pagesize=10
nextpage=Request.form("nextpage")
if nextpage="" then
  session("abspage")=1
else
  if nextpage="prev" then
    session("abspage")=session("abspage")-1
  elseif nextpage="next" then
    session("abspage")=session("abspage")+1
  elseif nextpage="first" then
    session("abspage")=1
  elseif nextpage="last" then
    session("abspage")=rs.pagecount
  end if
  rs.absolutepage=session("abspage")
end if
  if rs.recordcount>0 then
    i=0
    Response.write "<table border=1 width=100%/>"
    response.write "<tr><td colspan=5 align=center>"
    response.write "Fellow number: " & rs.recordcount
%>
  <tr>
    <td align=center width=25%>Date</td>
    <td align=center width=49%>Title</td>
    <td align=center width=10%>Author</td>
    <td align=center width=8%>Read</td>
    <td align=center width=8%>Fellow</td>
  </tr>
<%
  do while not rs.eof and i<10
%>
  <tr>
    <td align=center><%=rs("articledate")%>&nbsp;<%=rs("articletime")%></td>
    <td align=center><%=rs("articletitle")%></td>
    <td align=center><%=rs("articleauthor")%></td>
    <td align=center><%=rs("articleaccessnumber")%></td>
    <td align=center><%=rs("articlefellownumber")%></td>
  </tr>

<% rs.movenext
  i=i+1
  loop
  response.write "</table></center>"
  response.write "<center><form action showlist.asp method=post>"
  if rs.pagecount>1 then
    if (session("abspage"))>1 then
      response.write "<input type=submit value=prev name=nextpage>"
    end if
    if (session("abspage"))<rs.pagecount then
      response.write "<input type=submit value=next name=nextpage>"
    end if
  end if
  response.write "</form>"

  end if
  rs.close
  set rs=nothing
%>

</tr>
</table>

<hr>
<form action=publisharticle.asp method=post>
  <input type=hidden name=articleid value=<%=request.querystring("id")%>>
  <table border=0>
    <tr>
      <td>Title:</td>
      <td><input type=text name=title size=60></td>
    </tr>
    <tr>
      <td colspan=2>Content:</td>
    </tr>
    <tr>
      <td colspan=2><textarea id=textarea1 name=content style="height:100px;width:500px"></textarea></td>
    </tr>
  </table>
  <center>
    <br>
    <input id=submit1 name=submit type=submit value=fellow>
    <input id=submit1 name=submit type=submit value=publish>
  </center>
</form>
</body>
</html>
