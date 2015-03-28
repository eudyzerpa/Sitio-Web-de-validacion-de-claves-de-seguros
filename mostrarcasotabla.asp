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
<FORM METHOD="Post" name="consultastatus" ACTION="consultastatus.asp">
    <div align="center">
      <p><input type="hidden" name="consulta2" value="true">
  </p></div>
    
  <center><strong><font color="#000080" face="Verdana, Arial, Helvetica, sans-serif"> 
  </font></strong>
  </center>
  <H4 align="center" class="style1">&nbsp;</H4>

	<div align="center">
       
    <H4 align="center" class="style1">&nbsp;</H4>
<div align="center"></div>
  </div>
</FORM>
<%
        openstr = "driver={Microsoft Access Driver (*.mdb)};" & _
        "dbq=" & Server.MapPath("dbxsaps.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr
    
	
        sql = " SELECT * FROM Personas;" 
			  
		 
         Set rs = Server.CreateObject("ADODB.Recordset")
         rs.Open sql, cn, 3, 3 

        
%>
<div id="Layer1" "style="position:absolute; width:605px; height:115px; z-index:1; left: 4px; top: 22px; background-color: #000000; layer-background-color: #000000; border: 1px none #000000; visibility: visible; overflow: scroll;"> 
  <table width="100%" border="1">
    <tr bgcolor="#009933" class="style1" > 
      <th width="116" height="27"> <div align="center" class="style9"><font color="#FFFFFF">Cedula</font></div></th>
      <th width="118"> <div align="center" class="style9"><font color="#FFFFFF">Nombre</font></div></th>
      <th width="220"> <div align="center" class="style9"> <font color="#FFFFFF">Apellido</font></div></th>
      <th width="174"> <div align="center" class="style9"><font color="#FFFFFF">Fecha 
          de Nacimiento</font></div></th>
      <th width="78"> <div align="center" class="style9"><font color="#FFFFFF">Sexo</font></div></th>
      <th width="192"> <div align="center" class="style9"><font color="#FFFFFF">Estado 
          Civil</font></div></th>
    </tr>
    <% Do while Not rs.EOF %>
    <tr class="style1"> 
      <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%= rs("Documento") %></font></div></td>
      <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%= rs("Nombre") %></font></div></td>
      <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%= rs("PrimerApellido") %></font> </div></td>
      <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%= rs("Fnacimiento") %></font></div></td>
      <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%= rs("Sexo") %> </font></div></td>
      <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%= rs("EdoCivil") %></font></div></td>
    </tr>
    <% rs.MoveNext
	   Loop
	 %>
  </table>
</div>
</BODY>
</HTML>
