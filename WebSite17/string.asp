<%
set conn=Server.CreateObject("ADODB.Connection")
conn.Open "pd1"



keywords=request.form("keyword")

    set rs=Server.CreateObject("ADODB.recordset")
    sql="SELECT * FROM plist where fname = '" & keywords & "' or lname = '" & keywords & "'"
    rs.Open sql, conn


    %>


<html>

<body>
      <div>
        <a href="Homepage.asp">Homepage</a>
    </div>
<form  method="post">

<input type="text" name="keyword" size="20" />
<input type="submit" value="search" />

</form>

<table border="1" width="100%">
 
  <tr>
  <%if keywords <> " " then
        for each x in rs.Fields
             response.write("<th>" & x.name & "</th>")
        next
    else
      response.write()
    end if
      
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