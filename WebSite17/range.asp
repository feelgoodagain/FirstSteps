<%
set conn=Server.CreateObject("ADODB.Connection")
conn.Open "pd1"


rytpe=request.form("rytpe")
keywords1=CInt(request.form("keyword1"))



keywords2=CInt(request.form("keyword2"))

if rytpe="Integer" Then
   sql="SELECT * FROM plist where ID between " & keywords1 & "and " & keywords2 

   set rs=Server.CreateObject("ADODB.Recordset")
   rs.Open sql,conn

else
   sql="SELECT * FROM plist where salary between " & keywords1 & "and " & keywords2 

   set rs=Server.CreateObject("ADODB.Recordset")
   rs.Open sql,conn

end if

    %>


<html>
<body>
      <div>
        <a href="Homepage.asp">Homepage</a>
    </div>
<form method="post">
range type:
    <select name="rytpe">
        <option value="Integer">Integer</option>
        <option value="Floating">Floating</option>
        
    </select>

    <br />


from:
    <input type="text" name="keyword1" size="20" />
to:
    <input type="text" name="keyword2" size="20" />
    <input type="submit" value="search" />

</form>

<table border="1" width="100%">
 
  <tr>
  <%
        for each x in rs.Fields
             response.write("<th>" & x.name & "</th>")
        next


      
      %>
  </tr>
  <%do until rs.EOF%>
    <tr>
    <%for each x in rs.Fields%>
      <td><%Response.Write(x.value)%></td>
    <%next
    rs.MoveNext%>
    </tr>
  <%loop
  rs.close
  conn.close%>
</table>
</body>
</html>