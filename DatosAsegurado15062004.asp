<HTML>
<HEAD>
<TITLE>SIstema INterconectado de Recepción e Envío de Datos</TITLE>
<style type="text/css">
<!--
.style1 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 12;
	color: #000080;
}
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

Dim TICKET
Dim Siglas


    if session("Cedula") <> "" then
        openstr = "driver={Microsoft Access Driver (*.mdb)};" & _
        "dbq=" & Server.MapPath("dbxsinred.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr


        sql1 = " SELECT Siglas " & _
              " FROM Clinicas " & _
              " WHERE Entidad = '" & session("Clinica") & "'" 

        	Set rs1 = Server.CreateObject("ADODB.Recordset")
        	rs1.Open sql1, cn, 3, 3 

                 Siglas = rs1("Siglas")
                 session("Siglas") = Siglas
                
     
        sql = " SELECT * " & _
              " FROM Asegurados " & _
              " WHERE Entidad = '" & session("Cedula") & "'" 

        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open sql, cn, 3, 3 

        if rs.eof then
		    response.redirect("mensaje000090.asp")
        else 
		 	if rs.fields("Status") <> "" Then
			
				response.Redirect("mensaje000092.asp")
			
			Else
		
				IF LEN(CSTR(Year(date)))=2 THEN
				TICKET=Year(date)
				ELSE
				TICKET=mid(Year(date),3,2)
				END IF
				IF LEN( CSTR(Month(date)))=1 THEN
				TICKET=TICKET & "0" & Month(date) 
				ELSE
				TICKET=TICKET & Month(date) 
				END IF

				IF LEN( CSTR(Day(date)))=1 THEN
				TICKET=TICKET & "0" & Day(date)
				ELSE
					TICKET=TICKET & Day(date) 
				END IF

				IF LEN( CSTR(Hour(Time)))=1 THEN
					TICKET=TICKET & "0" & Hour(Time)
				ELSE
				TICKET=TICKET & Hour(Time)
				END IF

				IF LEN( CSTR(Minute(time)))=1 THEN
				TICKET=TICKET & "0" & Minute(time)
				ELSE
					TICKET=TICKET & Minute(time)
				END IF

				IF LEN( CSTR(Second(time)))=1 THEN
					TICKET=TICKET & "0" & Second(time)
				ELSE
				TICKET=TICKET & Second(time)
				END IF

					    

		    	sql = ""
			Sql  = "Insert Into CasosAbiertos "
			sql = sql & " ( "
			Sql = Sql & " Entidad,"
			Sql = Sql & " SubEntidad,"
			Sql = Sql & " Clinica,"
			Sql = Sql & " Poliza,"
			Sql = Sql & " TipodeApertura,"
			Sql = Sql & " Ticket,"
			Sql = Sql & " CartaAval,"
			Sql = Sql & " Nota,"
			Sql = Sql & " Anexo,"
			Sql = Sql & " Imagen,"
			Sql = Sql & " MinutosAutorizados,"
			Sql = Sql & " Recaudos,"
			Sql = Sql & " Diagnostico,"
			Sql = Sql & " MedicoTratante,"
			Sql = Sql & " Presupuesto,"
			Sql = Sql & " Autorizado,"
			Sql = Sql & " Usuario,"
			Sql = Sql & " FApertura"
	
			Sql = Sql & " ) "
			Sql = Sql & " Values "
			Sql = Sql & " ("
		
			Sql = Sql & "'" & session("Cedula") & "',"
		 	Sql = Sql & "'',"
			Sql = Sql & "'" & session("Clinica") & "',"
		 	Sql = Sql & "'" & rs("poliza") & "',"
			Sql = Sql & "'EMERGENCIA',"
			Sql = Sql & "'" & Siglas & TICKET & "',"
			Sql = Sql & "'',"
			Sql = Sql & "0,"
			Sql = Sql & "0,"
			Sql = Sql & "0,"
			Sql = Sql & "'" & rs("MinutosAutorizados") & "',"
    		Sql = Sql & "'NO',"
			Sql = Sql & "'',"
			Sql = Sql & "'',"
		   	Sql = Sql & "0,"
			Sql = Sql & "'NO',"
			Sql = Sql & "'" & session("Usuario") & "',"
			Sql = Sql & "#" & Now & "#"
			Sql = Sql & ")"
		
		
		cn.execute Sql
	'	session("ClavedeIngreso") = Siglas & TICKET

              	
		
		End if


             
		
	     sql2 = " SELECT * " & _
                " FROM Personas " & _
                " WHERE Entidad = '" & session("Cedula") & "'" 
			  
		 Set rsx = Server.CreateObject("ADODB.Recordset")
         rsx.Open sql2, cn, 3, 3 
		 
	       
			 	 
		
		%>
		<div id="Layer1" style="position:absolute; width:340px; height:24px; z-index:7; left: 9px; top: 1px;"> 
  <table width="101%" border="0">
    <tr> 
      <td><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <%response.write "<STRONG>" & "<CENTER>" & "DATOS DEL ASEGURADO" & "</CENTER>" & "</STRONG>" 
	 response.write "<CENTER>" & "ASEGURADO Y POLIZA VIGENTE, CUMPLE TODOS LOS REQUISITOS" & "</CENTER>" 
	 response.write "<CENTER>" & "SERVICIO GARANTIZADO" & "</CENTER>" & "<br>"
	 response.write "<STRONG>" &"<CENTER>" & "TICKET: "&  SiGLAS  & TICKET & "</CENTER>" & "</STRONG>" & "<br>"%>
        </font></td>
    </tr>
  </table>
</div>
<div id="Layer2" style="position:absolute; left:47px; top:84px; width:302px; height:106px; z-index:6"> 
  <table width="100%" border="0">
    <tr> 
      <td width="31%"><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif">Nombre:</font></td>
      <td width="69%"><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif"><% =rsx.Fields("Nombre")%> 
        </font></td>
    </tr>
    <tr> 
      <td><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif">Apellidos:</font></td>
      <td><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%  =rsx.Fields("PrimerApellido") & " " & rsx.Fields("SegundoApellido")%> 
        </font></td>
    </tr>
    <tr> 
      <td><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cedula:</font></td>
      <td><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%  =rsx.Fields("Documento")%> 
        </font></td>
    </tr>
    <tr> 
      <td><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif">FNacimiento:</font></td>
      <td><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%  =rsx.Fields("FNacimiento")%> 
        </font></td>
    </tr>
    <tr> 
      <td><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif">Sexo:</font></td>
      <td><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%  =rsx.Fields("Sexo")%> 
        </font></td>
    </tr>
    <tr> 
      <td><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif">Estado 
        Civil:</font></td>
      <td><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif"><% =rsx.Fields("EstadoCivil")%> 
        </font></td>
    </tr>
    <tr> 
      <td><font color="#FF0000" size="1" face="Verdana, Arial, Helvetica, sans-serif">COBERTURA</font></td>
      <td><font color="#FF0000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Bs. 
        30.000.000</font></td>
    </tr>
    <tr> 
      <td><font color="#FF0000" size="1" face="Verdana, Arial, Helvetica, sans-serif">DEDUCIBLE</font></td>
      <td><font color="#FF0000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Bs. 
        250.000</font></td>
    </tr>
  </table>
</div>

<% end if

		rs.Close
		Set rs = Nothing
		
		cn.Close
		Set cn = Nothing

END IF
%>
<div id="Layer4" style="position:absolute; left:313px; top:213px; width:74px; height:28px; z-index:5"> 
  <a href="sesion.asp"><img src="botonmenu.gif" width="57" height="28" border="0"></a> 
</div>
<p align="center"><strong><font color="#000080" size="2"> </font></strong> </p>
</BODY>
</HTML>
