<%


set conn=Server.CreateObject("ADODB.Connection")
conn.Open "pd1"

keyword1=request.form("keyword1")
keyword2=request.form("keyword2")
if isdate(keyword1) then
keywords1=DateValue(keyword1)
keywords2=DateValue(keyword2)




   sql="SELECT * FROM plist where ID between 1 and 45"
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



from:
    <input type="text" name="keyword1" size="20" />

to:
    <input type="text" name="keyword2" size="20" />


    <input type="submit" value="search" />

</form>

<table border="1" width="100%">
 
  <tr>
  <%if keywords1<>"" then
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
  conn.close
      end if%>
</table>

</body>
</html>