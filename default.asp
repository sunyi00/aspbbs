<html>
<body bgcolor=#ffffff>
<br><br>
<p>
<dl>
  <dd>
    <div align=center style="font-size:xx-large">
      <font color="crimson">
        <p>BBS</p>
      </font>
    </div>
  </dd>
</dl>
<br><br>
<p align=center style="font-size:xx-large">
  <font size=4>
    welcome to BBS
  </font>
</p>

<center>
you are visitor No.
<%
  dim visitors
    whichfile=server.mappath("counter/counter.txt")
    set fs=server.createobject("Scripting.FileSystemObject")
    set thisfile=fs.opentextfile(whichfile)
    visitors=thisfile.readline
    thisfile.close
    countlen=len(visitors)
  Response.Write visitors
  visitors=visitors+1
  set out=fs.createtextfile(whichfile)
  out.writeline(visitors)
  out.close
  set fs=nothing
%>
</center>

<hr>

<table align=center border=0 width=630>
  <tr>
    <td width=20% align=center><a href="register.htm">Register</a></td>
    <td width=20% align=center><a href="login.htm">Login</a></td>
    <td width=20% align=center><a href="visitor.htm">Guest</a></td>
    <td width=20% align=center><a href="help.htm">Help</a></td>
    <td width=20% align=center><a href="contact.htm">Contact</a></td>
  </tr>
</table>

</body>
</html>
