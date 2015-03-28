<%

if session("LoggedIn") = 0 then
       response.redirect("login.asp?autorizado=falso")
    End if

    if request.form("consulta") = "true" then
        openstr = "driver={Microsoft Access Driver (*.mdb)};" & _
        "dbq=" & Server.MapPath("dbxsinred.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr
        
   'Abre un recordset de Casos abiertos para evaluar el campo diagnostico
    Sql = ""
    Sql = "Select * " &_
          "From casosabiertos " &_
          "Where Ticket = '" & request.form("Ticket") & "'"
          
    
    Set rs = server.CreateObject("ADODB.Recordset")
    rs.Open Sql,cn, 3, 3
    


    If rs.eof then
         response.Redirect("mensaje000012.asp")         
    Else
      	xEstatus = rs.fields("Estatus")
        xmontopresupuesto = rs.fields("Presupuesto")   
   		xdiagnostico = rs.fields("Diagnostico")
       
   if xdiagnostico <> "" Then 
   'Crear el Caso por Liquidar
    
		Sql = ""
		Sql = Sql & " INSERT INTO CasosPorLiquidar ("
		Sql = Sql & " Entidad,"
		Sql = Sql & " SubEntidad,"
		Sql = Sql & " Clinica,"
		Sql = Sql & " Poliza,"
		Sql = Sql & " Certificado,"
		Sql = Sql & " Riesgo,"
		Sql = Sql & " Parentesco,"
		Sql = Sql & " Ticket,"
		Sql = Sql & " TipodeApertura,"
		Sql = Sql & " CartaAval,"
		Sql = Sql & " FApertura,"
		Sql = Sql & " MinutosAutorizados,"
		Sql = Sql & " Nota,"
		Sql = Sql & " Anexo,"
		Sql = Sql & " Imagen,"
		Sql = Sql & " MedicoTratante,"
		Sql = Sql & " Baremo,"
		Sql = Sql & " Diagnostico,"
		Sql = Sql & " Presupuesto,"
		Sql = Sql & " FSolicitudDeAutorizacion,"
		Sql = Sql & " Motivo,"
		Sql = Sql & " Recaudos,"
		Sql = Sql & " Autorizado,"
		Sql = Sql & " MontoAutorizado,"
		Sql = Sql & " FAutorizacion,"
		Sql = Sql & " FSolicitudDeLiquidacion,"
		Sql = Sql & " Montoliquidado,"
		Sql = Sql & " Fliquidacion,"
		Sql = Sql & " TipodeCierre,"
		Sql = Sql & " FCierre,"
		Sql = Sql & " Siniestro,"
		Sql = Sql & " MontoPagado,"
		Sql = Sql & " FPago,"
		Sql = Sql & " ReferenciaPago,"
		Sql = Sql & " Usuario,"
		Sql = Sql & " Supervisor,"
		Sql = Sql & " Responsable,"
		Sql = Sql & " Medico,"
		Sql = Sql & " Observaciones"
	    
		Sql = Sql & " )"
	    
		Sql = Sql & " SELECT "
		Sql = Sql & " Entidad,"
		Sql = Sql & " SubEntidad,"
		Sql = Sql & " Clinica,"
		Sql = Sql & " Poliza,"
		Sql = Sql & " Certificado,"
		Sql = Sql & " Riesgo,"
		Sql = Sql & " Parentesco,"
		Sql = Sql & " Ticket,"
		Sql = Sql & " TipodeApertura,"
		Sql = Sql & " CartaAval,"
		Sql = Sql & " FApertura,"
		Sql = Sql & " MinutosAutorizados,"
		Sql = Sql & " Nota,"
		Sql = Sql & " Anexo,"
		Sql = Sql & " Imagen,"
		Sql = Sql & " MedicoTratante,"
		Sql = Sql & " Baremo,"
		Sql = Sql & " Diagnostico,"
		Sql = Sql & " Presupuesto,"
		Sql = Sql & " FSolicitudDeAutorizacion,"
		Sql = Sql & " Motivo,"
		Sql = Sql & " Recaudos,"
		Sql = Sql & " Autorizado,"
		Sql = Sql & " MontoAutorizado,"
		Sql = Sql & " FAutorizacion,"
		Sql = Sql & " FSolicitudDeLiquidacion,"
		Sql = Sql & " Montoliquidado,"
		Sql = Sql & " Fliquidacion,"
		Sql = Sql & " TipodeCierre,"
		Sql = Sql & " FCierre,"
		Sql = Sql & " Siniestro,"
		Sql = Sql & " MontoPagado,"
		Sql = Sql & " FPago,"
		Sql = Sql & " ReferenciaPago,"
		Sql = Sql & " Usuario,"
		Sql = Sql & " Supervisor,"
		Sql = Sql & " Responsable,"
		Sql = Sql & " Medico,"
		Sql = Sql & " Observaciones"
		Sql = Sql & " FROM "
		Sql = Sql & " CasosAbiertos"
		Sql = Sql & " WHERE "
		Sql = Sql & " Ticket  ='" & request.form("Ticket") & "'" 
        
		cn.Execute Sql
    
		'Actualizar Casos por Liquidar
    
		usuario = session("usuario")

	if xEstatus <> "ABIERTO" Then
    
    	Sql = ""
   		Sql = Sql & " UPDATE CasosPorLiquidar SET"
   		Sql = Sql & " TipoDeCierre = 'LIQUIDACION',"
   		Sql = Sql & " Estatus = 'POR LIQUIDAR',"
   		Sql = Sql & " FSolicitudDeLiquidacion = #" & Now & "#,"
   		Sql = Sql & " Usuario = '" & Usuario & "'"
   		Sql = Sql & " WHERE Ticket  ='" & request.form("Ticket") & "'"     
      
	Else

   		Sql = "UPDATE "
   		Sql = Sql & " CasosPorLiquidar SET"
		Sql = Sql & " TipoDeCierre = 'LIQUIDACION',"
   		Sql = Sql & " Estatus = 'POR LIQUIDAR',"
   		Sql = Sql & " FSolicitudDeLiquidacion = #" & Now & "#,"
   		Sql = Sql & " Usuario = '" & Usuario & "',"
   		Sql = Sql & " Autorizado = 'AC',"
  		Sql = Sql & " FSolicituddeAutorizacion = #" & Now & "#,"
  		Sql = Sql & " FAutorizacion = #" & Now & "#,"
   		Sql = Sql & " OBSAutorizacion = 'AUTORIZACION ASUMIDA POR LA CLINICA',"
  		Sql = Sql & " MontoAutorizado = " & xmontopresupuesto & ""
   		Sql = Sql & " WHERE "
   		Sql = Sql & " Ticket = '" & request.form("Ticket") & "'" 
         
	END IF

    cn.Execute Sql
    
    'Eliminar el caso Abierto
    
    Sql = ""
    Sql = Sql & " DELETE "
    Sql = Sql & " FROM "
    Sql = Sql & " CasosAbiertos"
    Sql = Sql & " WHERE Ticket  ='" & request.form("Ticket") & "'" 
    
						
	 cn.Execute sql, raffected
         
         if raffected > 0 then
              response.Redirect("mensaje000096.asp")
         else
              response.Redirect("mensaje000070.asp")
         END IF
         
    ELSE
		Response.Redirect ("mensaje001000")   	
    END IF
    
   END IF				
 END IF					
	
%>
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
.style3 {font-family: Verdana, Arial, Helvetica, sans-serif}
.style4 {font-size: 12}
.style5 {color: #000080}
.style6 {font-family: Verdana, Arial, Helvetica, sans-serif; color: #000080; }
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
<BODY Background="back2.jpg" vlink="black" link="black">
<p>&nbsp;</p>
<p>&nbsp;</p>
<FORM METHOD="Post" name="ClaveIngreso" ACTION="clavedeingresocerrarcasos.asp">
    <div align="center">
      <input type="hidden" name="consulta" value="true">
  </div>
    
  <H4 align="center" class="style1">&nbsp; </H4>
	<div align="center">
	  
    <TABLE width="279" BORDER=0>
      <TR>
        <TD class="style4 style5">TICKET
        <TD class="style4"><INPUT NAME="Ticket" SIZE="15"> 
      <TR>
        <TD COLSPAN=2 class="style4"><input name="Submit" type="Submit" value="Enviar">	
    </TABLE>
  </div>
</FORM>
</BODY>
</HTML>

