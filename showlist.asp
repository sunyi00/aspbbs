<%@language=vbscript%>
<!--#include file=conn.asp-->
<html>
<head>
  <title>Article List</title>
</head>
<body bgcolor=#ffffff>
<p align=center><font size=5>article list</p>

<%
  dim strselectsql
  dim intarticleid
  if request.querystring("id")="" then
    intarticleid=0
  end if
  set rs=server.createobject("adodb.recordset")
  strselectsql="select * from article where articleparent=" & intarticleid
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
    Response.Write "<table border=1 width=100%>"
    Response.Write "<tr><td colspan=5 align=center>"
    if intarticleid=0 then
      Response.Write "Page " & session("abspage") & "&nbsp; Totally " & rs.RecordCount & " topics"
    else
      Response.Write "fellows number: " & rs.RecordCount
    end if
    Response.Write "</tr>"
    %>
      <tr>
        <td align=center width=19%>Date</td>
        <td align=center width=55%>Title</td>
        <td align=center width=10%>Author</td>
        <td align=center width=8%>Read</td>
        <td align=center width=8%>Fellows</td>
      </tr>
    <%
    do while not rs.eof and i<10
    %>
    <tr>
      <td align=center><%=rs("articledate")%>&nbsp;<%=rs("articletime")%></td>
      <td><a href=readarticle.asp?id=<%=rs("articleid")%>><%=rs("articletitle")%></a></td>
      <%if session("name")="" then%>
        <td align=center><%=rs("articleauthor")%></td>
      <%else%>
        <td align=center><a href=newmessage.asp?toname=<%=rs("articleauthor")%>><%=rs("articleauthor")%></td>
      <%end if%>
      <td align=center><%=rs("articleaccessnumber")%></td>
      <td align=center><%=rs("articlefellownumber")%></td>
    </tr>
    <%rs.movenext
        i=i+1
        loop
        Response.Write "</table></center>"
        Response.Write "<center><form action showlist.asp method=post>"
        if rs.pagecount>1 then
          Response.Write "<input type=submit value=prev name=nextpage> & " "
        end if
        if session("abspage") < rs.pagecount then
          Response.Write "<input type=submit value=next name=nextpage>"
        end if
        Response.Write "</form>"
  end if
  rs.close
  set rs=nothing
%>
  </tr>
  </table>
</body>
</html>
