<%
set conn=Server.CreateObject("ADODB.Connection")
conn.Open "pd1"



keywords1=request.form("keyword1")



keywords2=CInt(request.form("keyword2"))



   sql="SELECT plist.fname, plist.lname FROM plist left outer join clist on plist.ID = clist.ID where clist.ctype= '" & keywords1 & "'"

   set rs=Server.CreateObject("ADODB.Recordset")
   rs.Open sql,conn



    %>


<html>
<body>
      <div>
        <a href="Homepage.asp">Homepage</a>
    </div>
<form method="post">

   

Search for who drives the car with type of:
    <input type="text" name="keyword1" size="20" />

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