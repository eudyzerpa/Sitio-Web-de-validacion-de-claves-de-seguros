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
    
		
    Sql = ""
    Sql = Sql & " Select "
    
    Sql = Sql & " Personas.Nombre,"
    Sql = Sql & " Personas.PrimerApellido,"
    Sql = Sql & " Personas.SegundoApellido,"
    Sql = Sql & " Personas.TipodeDocumento,"
    Sql = Sql & " Personas.Documento"
        
    Sql = Sql & " From "
    Sql = Sql & " CasosAbiertos "
    
    Sql = Sql & " Left Join Personas on CasosAbiertos.Entidad = Personas.Entidad"
    Sql = Sql & " WHERE CasosAbiertos.Ticket = '" & session("ClavedeIngreso") & "'"
		 
			  
		 
         Set rs = Server.CreateObject("ADODB.Recordset")
         rs.Open sql, cn, 3, 3 
		 
		 if Not rs.EOF Then
		 	xNombre = rs.Fields("Nombre")
			xApellido = rs.fields("PrimerApellido")		 
		  	xCedula  = rs.fields("Documento")
		 End If
		 
		 Session("Medico")= request.Form("Txt_Medico")
		 Session("Diagnostico")= request.Form("Txt_Diagnostico")
		 
		 rs.close
		 Set rs = Nothing        
		 
		 cn.close
		 set cn = nothing

 
  %>
<table width="100%" border="0">
  <tr> 
    <td width="11%"><strong><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif">Clave:</font></strong></td>
    <td width="89%"> <font color="#000099" size="1">
      <% response.Write(session("ClavedeIngreso"))  %>
      </font> </td>
  </tr>
  <tr> 
    <td><strong><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif">Asegurado</font></strong></td>
    <td><font color="#000099" size="1">
      <% = xNombre & " " & xapellido %>
      </font> </td>
  </tr>
  <tr> 
    <td height="20"><strong><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cedula:</font></strong></td>
    <td><font color="#000099" size="1">
      <% = xcedula %>
      </font></td>
  </tr>
</table>
<%

if Request.Form("comp") = "actualizar" then

        openstr = "driver={Microsoft Access Driver (*.mdb)};" & _
        "dbq=" & Server.MapPath("dbxsaps.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr


 '   SQL = "UPDATE CasosAbiertos set MedicoTratante = '" & request.Form("Txt_Medico") & "'," & _
	'      " Diagnostico = '" & request.Form("Txt_Diagnostico") & "'" & _
	 '     " WHERE CasosAbiertos.ClavedeIngreso = '" & session("ClavedeIngreso") & "'"
	
	 sql="SELECT * FROM CasosAbiertos WHERE CasosAbiertos.ClavedeIngreso = '" & session("ClavedeIngreso") & "'"
	
	     Set rs = Server.CreateObject("ADODB.Recordset")
         rs.Open sql
		 
		 on error resume next
		 
		 if Not rs.EOF Then
		 
		 	rs.Fields("MedicoTratante")="HOLA"
			rs.fields("Diagnostico")="HOLA"

			response.Write(SQL)	                                
		 	Response.Write "Actualizando"		 		 
		 	rs.update		 
			
		 	if err.number<>0 then
				response.write "Error " & err.description
			on error goto 0
			else
				on error goto 0
			end if
	
	end if
	  
	 cn.Execute SQL                 
     response.Write(SQL)
	 
	' on error resume next
     'cn.Execute SQL    
	 
	' if err.number<>0 then
	 '	response.write "Error " & err.description
	 '	on error goto 0
	'else
	'	on error goto 0
	'end if
	
	 'rs.close
	 'SET rs = Nothing
	 
	 'n.close
	 'set cn = nothing
	 
End if
%>	  

<div id="Layer1" style="position:absolute; width:235px; height:165px; z-index:1; left: 11px; top: 70px;"> 
  <form name="Actualizar" method="post" action="actualizardiagnosticodelcaso.asp">
    <table width="76%" border="0" >
      <tr> 
        <td width="35%" height="24"><strong><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif">Fecha</font></strong></td>
        <td width="65%"> <input type="text" size="12" name="Txt_Medico"></td>
      </tr>
      <tr> 
        <td height="24"><strong><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif">Referencia</font></strong></td>
        <td><input type="text" size="12" name="Txt_Medico4"> </td>
      </tr>
      <tr> 
        <td height="24"><strong><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cargo</font></strong></td>
        <td><input type="text" size="12" name="Txt_Diagnostico"></td>
      </tr>
      <tr> 
        <td height="24"><strong><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cantidad</font></strong></td>
        <td><input type="text" size="12" name="Txt_Medico2"></td>
      </tr>
      <tr> 
        <td height="24"><strong><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif">Obs</font></strong></td>
        <td><input type="text" size="12" name="Txt_Medico3"></td>
      </tr>
      <tr>
        <td height="25"><a href="mensaje000094.asp"><img src="enviar.gif" width="57" height="28" border="0"></a></td>
        <td><input type="hidden" name="comp" value="actualizar">
        </td>
      </tr>
    </table>
    </form>
</div>
</BODY>
</HTML>
