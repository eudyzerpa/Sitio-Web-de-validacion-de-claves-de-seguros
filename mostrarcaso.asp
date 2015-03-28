<%
        ' ABRE LA CONEXION CON LA BASE DE DATOS
        openstr = "driver={Microsoft Access Driver (*.mdb)};" & _
        "dbq=" & Server.MapPath("dbxsaps.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr

        sql = " SELECT * " & _
              " FROM Personasx "
        'ABRE EL RECORDSET
		
	    Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open sql, cn, 3, 3 
		
 if rs.EOF and rs.bof then
        %>
            <script Language="JavaScript">
                alert('No existen registros que cumplan con las condiciones de búsqueda de esta consulta');
            </script>
        <%
  	else
		rs.MoveFirst
       
		while not rs.EOF 

%>
		 <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr bgcolor="<%=BgColor%>"> 
            <td height="18" width="27%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">&nbsp;&nbsp;<%=Entidad%></font></td>
            <td height="18" width="23%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%=TipodeDocumento%></font></td>
            <td height="18" width="16%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%=Documento%></font></td>
            <td height="18" width="34%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%=Nombre%></font></td>
          </tr>
        </table>
<%	

            rs.MoveNext 
       		wend
%>