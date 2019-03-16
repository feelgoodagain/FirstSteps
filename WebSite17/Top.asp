<%
set conn=Server.CreateObject("ADODB.Connection")
conn.Open "pd1"


rytpe=request.form("rytpe")
keywords1=CInt(request.form("keyword1"))





if rytpe="Integer" and keywords1>0 Then
    
   sql="SELECT TOP  " & keywords1 & " ID FROM plist ORDER by ID ASC"

   set rs=Server.CreateObject("ADODB.Recordset")
   rs.Open sql,conn

elseif rytpe="Floating" and keywords1>0 Then
   sql="SELECT TOP  " & keywords1 & " salary FROM plist ORDER by salary ASC "

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
data type:
    <select name="rytpe">
   
        <option value="Integer">Integer</option>
        <option value="Floating">Floating</option>
        
    </select>

    <br />


         search for the lowest
    <input type="text" name="keyword1" size="20" />
    information
    
    <input type="submit" value="search" />

</form>

<table border="1" width="100%">
 
  <tr>
  <%
      if keywords1>0 then
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