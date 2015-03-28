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
    Sql = Sql & " CasosAbiertos.Entidad,"
        
    
    Sql = Sql & " CasosAbiertos.Clinica,"
    Sql = Sql & " CasosAbiertos.Poliza,"
    
    Sql = Sql & " Personas.Nombre,"
    Sql = Sql & " Personas.PrimerApellido,"
    Sql = Sql & " Personas.SegundoApellido,"
    Sql = Sql & " Personas.TipodeDocumento,"
    Sql = Sql & " Personas.Documento,"

    Sql = Sql & " Personas.FNacimiento,"
    Sql = Sql & " Personas.Sexo,"
    Sql = Sql & " Personas.EstadoCivil,"
    
    Sql = Sql & " Asegurados.DescCategoria as [Plan ó Categoria],"
    Sql = Sql & " CobxCert.Certificado,"
    Sql = Sql & " CobxCert.Cobertura,"
    Sql = Sql & " CobxCert.Monto,"
    Sql = Sql & " CobxCert.Deducible,"
    Sql = Sql & " CobxCert.CoberturaConsumida,"
    Sql = Sql & " CobxCert.Exceso,"
        
    
    Sql = Sql & " CasosAbiertos.TipodeApertura,"
    Sql = Sql & " CasosAbiertos.Ticket as Ticket,"
    Sql = Sql & " CasosAbiertos.CartaAval,"
    Sql = Sql & " CasosAbiertos.Diagnostico,"
    Sql = Sql & " CasosAbiertos.MedicoTratante as Medico,"
    Sql = Sql & " CasosAbiertos.Usuario"
    
        
    Sql = Sql & " From "
    Sql = Sql & " (( "
    Sql = Sql & " CasosAbiertos"
    
    Sql = Sql & " Left Join Personas on CasosAbiertos.Entidad = Personas.Entidad"
    Sql = Sql & " )"
    Sql = Sql & " Left Join Asegurados on CasosAbiertos.Entidad = Asegurados.Entidad"
    Sql = Sql & " )"
    Sql = Sql & " Left Join CoberturaXCertificado as CobxCert on CasosAbiertos.Entidad = CobxCert.Entidad"
   'Sql = Sql & " WHERE CLINICA = '" & session("Clinica") & "'"
		 
			  
		 
         Set rs = Server.CreateObject("ADODB.Recordset")
         rs.Open sql, cn, 3, 3 

        
%>
<H4 align="center"><font color="#000080" face="Verdana, Arial, Helvetica, sans-serif"><!-- <% response.Write(session("Clinica")) %> --></font></h4>
<div id="Layer1" "style="position:absolute; width:845px; height:98px; z-index:1; left: 13px; top: 82px; background-color: #000000; layer-background-color: #000000; border: 1px none #000000; visibility: visible; overflow: scroll;"> 
  <table width="100%" border="1">
    <tr bgcolor="#009933" class="style1" > 
      <th width="95" height="27"> <div align="center" class="style9"><font color="#FFFFFF">Ticket</font></div></th>
      <th width="95"> <div align="center" class="style9"><font color="#FFFFFF">Cedula</font></div></th>
      <th width="95"> <div align="center" class="style9"><font color="#FFFFFF">Nombre</font> 
        </div></th>
      <th width="95"> <div align="center" class="style9"><font color="#FFFFFF">Apellido</font></div></th>
      <th width="95"> <div align="center" class="style9"><font color="#FFFFFF">FNacimiento</font></div></th>
      <th width="95"> <div align="center" class="style9"><font color="#FFFFFF">Sexo</font></div></th>
      <th width="95"><div align="center"><font color="#FFFFFF">Edo. Civil</font></div></th>
      <th width="95"><font color="#FFFFFF">Diagnostico</font></th>
      <th width="95"><font color="#FFFFFF">Apertura</font></th>
      <th width="200"><font color="#FFFFFF">&nbspFInicio</font></th>
    </tr>
    <% if Not rs.EOF then %>
    <% Do while Not rs.EOF %>
    <tr class="style1"> 
      <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%= rs("Ticket") %></font></div></td>
      <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%= rs("Documento") %></font></div></td>
      <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%= rs("Nombre") %></font> 
        </div></td>
      <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%= rs("PrimerApellido") %></font></div></td>
      <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%= rs("Fnacimiento") %> 
          </font></div></td>
      <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%= rs("Sexo") %></font></div></td>
      <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%= rs("EstadoCivil") %></font></div></td>
      <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%= rs("Diagnostico") %></font></div></td>
      <td><div align="center">
        <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%= rs("TipodeApertura") %></font></div></td>
      <td width="200"><div align="center">
        <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%= rs("TipodeApertura") %></font></div></td>
    </tr>
    <% rs.MoveNext
	   Loop
	 %>
    <% Else 
	    response.Write("No se han encontrado casos")
	 end if
	 %>
  </table>
</div>
</BODY>
</HTML>
