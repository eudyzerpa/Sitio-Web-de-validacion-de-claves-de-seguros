<HTML>
<HEAD>
<TITLE>Active Server Pages</TITLE>
<style type="text/css">
<!--
.style1 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 12;
	color: #000080;
}
.style9 {color: #000099}
-->
</style>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
</script>
</HEAD>
<BODY BGCOLOR=FFFFFF>
<%
        openstr = "driver={Microsoft Access Driver (*.mdb)};" & _
        "dbq=" & Server.MapPath("dbxsinred.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr
    
	
        'sql = " SELECT * FROM CasosAbiertos WHERE Clinica = 'CLINICAS CARACAS'" 
		
    Sql = ""
    Sql = Sql & " Select "
    Sql = Sql & " Cargos.Id, "
    Sql = Sql & " Cargos.Entidad,"
    Sql = Sql & " Cargos.SubEntidad,"
    
       
    Sql = Sql & " Cargos.Ticket,"
    Sql = Sql & " Cargos.Eliminado,"
    Sql = Sql & " Cargos.Fecha,"
    Sql = Sql & " Cargos.Descripcion,"
    Sql = Sql & " Cargos.Cantidad,"
    Sql = Sql & " Cargos.Costo,"
    Sql = Sql & " Cargos.Usuario,"
    Sql = Sql & " Cargos.Status,"
	Sql = Sql & " Cargos.Referencia,"
    Sql = Sql & " Cargos.Observaciones1,"
    Sql = Sql & " Cargos.Observaciones2,"
    Sql = Sql & " Cargos.FRegistro"
    
    Sql = Sql & " From "
    Sql = Sql & " Cargos"
    Sql = Sql & " WHERE Cargos.Ticket = '" & session("ClaveDeIngreso") & "'"
		 
			  
		 
         Set rs = Server.CreateObject("ADODB.Recordset")
         rs.Open sql, cn, 3, 3 

 
        
%>
<H4 align="center"><font color="#000080" face="Verdana, Arial, Helvetica, sans-serif"></font></h4>
<div id="Layer1" "style="position:absolute; width:550px; height:98px; z-index:1; left: -9px; top: 82px; background-color: #000000; layer-background-color: #000000; border: 1px none #000000; visibility: visible; overflow: scroll;"> 
  <table width="109%" border="1" align="center">
    <tr bgcolor="#009933" class="style1" > 
      <th width="100"> <div align="center" class="style9"><font color="#FFFFFF">Fecha</font></div></th>
      <th width="100"> <div align="center" class="style9"><font color="#FFFFFF">Referencia</font> 
        </div></th>
      <th width="100"> <div align="center" class="style9"><font color="#FFFFFF">Descripci&oacute;n</font></div></th>
      <th width="100"> <div align="center" class="style9"><font color="#FFFFFF">Cantidad</font></div></th>
      <th width="100"> <div align="center" class="style9"><font color="#FFFFFF">Costo</font></div></th>
      <th width="100"><div align="center" class="style9"><font color="#FFFFFF">Fecha</font></th>
      <th width="100"><div align="center" class="style9"><font color="#FFFFFF">Status</font></th>
      <th width="100"><div align="center" class="style9"><font color="#FFFFFF">Observaciones</font></th>
    </tr>
    <% if Not rs.EOF then %>
    <% Do while Not rs.EOF %>
    <tr class="style1"> 
      <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%= rs("Fecha") %></font></div></td>
      <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%= rs("Referencia") %></font> 
        </div></td>
      <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%= rs("Descripcion") %></font></div></td>
      <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%= rs("Cantidad") %> 
          </font></div></td>
      <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%= rs("Costo") %></font></div></td>
      <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%= rs("Fecha") %></font></div></td>
      <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%= rs("Status") %></font></div></td>
      <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%= rs("Descripcion") %></font></div></td>
    </tr>
    <% rs.MoveNext
	   Loop
	 %>
    <% Else 
	    response.redirect("mensaje000033.asp")
	 end if
	 %>
  </table>
</div>
</BODY>
</HTML>
