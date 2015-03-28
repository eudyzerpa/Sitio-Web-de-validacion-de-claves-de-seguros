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
		 
		 
   
       if Request.Form("comp") = "actualizar" then
	   
	   
	   		sql = ""
			Sql  = "Insert Into Cargos "
			sql = sql & " ( "
		
        	Sql = Sql & " Descripcion,"
		'	Sql = Sql & " Cantidad,"
		'	Sql = Sql & " Costo,"
			Sql = Sql & " Referencia,"
		'	Sql = Sql & " Usuario,"
		'	Sql = Sql & " Status,"
			Sql = Sql & " Observaciones1"
		'	Sql = Sql & " Observaciones2"
				
			Sql = Sql & " ) "
			Sql = Sql & " Values "
			Sql = Sql & " ("
		
		'	Sql = Sql & "'" & session("ClavedeIngreso") & "',"
		' 	Sql = Sql & "'" & session("ClavedeIngreso") & "',"
		'	Sql = Sql & "'00000001',"
		'	Sql = Sql & "'" & session("Clinica") & "',"
		 '	Sql = Sql & "'" & rs("poliza") & "',"
		'	Sql = Sql & "'" & Request.Form("Txt_Fecha")  & "',"
			Sql = Sql & "'" & Request.Form("Txt_Descripcion")  & "',"
		'' 	Sql = Sql & "'" & Request.Form("Txt_Cantidad")  & "',"
		'	Sql = Sql & "'',"
			Sql = Sql & "'" & Request.Form("Txt_Referencia")  & "',"
		'	Sql = Sql & "0,"
		'	Sql = Sql & "0,"
			Sql = Sql & "'" & Request.Form("Txt_Observaciones1")  & "'"
       	'	Sql = Sql & "'NO',"
			Sql = Sql & ")"
				
		cn.execute Sql
 End if
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



      

<div id="Layer1" style="position:absolute; width:450px; height:115px; z-index:1; left: 10px; top: 99px;"> 
  <form name="Actualizar" method="post" action="cargosnotarifados.asp">
    <table width="45%" border="0" >
      <tr> 
        <td width="46%" height="24"><strong><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif">Fecha</font></strong></td>
        <td width="54%"><input type="text" size="12" name="Txt_Fecha"> </td>
      </tr>
      <tr> 
        <td height="24"><strong><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif">Referencia</font></strong></td>
        <td><input type="text" size="12" name="Txt_Referencia"> </td>
      </tr>
      <tr> 
        <td height="24"><strong><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif">Descripcion</font></strong></td>
        <td><input type="text" size="12" name="Txt_Descripcion"></td>
      </tr>
      <tr> 
        <td height="24"><strong><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cantidad</font></strong></td>
        <td><input type="text" size="12" name="Txt_Cantidad"></td>
      </tr>
      <tr> 
        <td height="24"><strong><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif">Costo</font></strong></td>
        <td><input type="text" size="12" name="Txt_Costo"></td>
      </tr>
      <tr> 
        <td height="24"><strong><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif">Obs</font></strong></td>
        <td><input type="text" size="12" name="Txt_Observaciones"></td>
      </tr>
      <tr>
        <td height="25"><a href="JavaScript:document.Actualizar.submit();"><img src="enviar.gif" width="57" height="28" border="0"></a></td>
        <td><input type="hidden" name="comp" value="actualizar">
        </td>
      </tr>
    </table>
    
    </form>
</div>

</BODY>
</HTML>
