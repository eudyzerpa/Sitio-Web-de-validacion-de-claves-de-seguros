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
<BODY Background="back2.jpg" vlink="black" link="black">
<strong> 
<%

Dim TICKET
Dim Siglas
Dim txt_cedula


    if session("Cedula") <> "" then
        openstr = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=rdmcds;Initial Catalog=DBXSINRED;Data Source=dell1600"
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
            Sql = Sql & " Responsable,"
			Sql = Sql & " Supervisor,"
			Sql = Sql & " Medico,"
			Sql = Sql & " Estatus,"
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
            Sql = Sql & "'',"
			Sql = Sql & "'',"
            Sql = Sql & "'',"
            Sql = Sql & "'ABIERTO',"
			Sql = Sql & "Getdate()"
			Sql = Sql & ")"
		
		
			cn.execute Sql
	
              	
		
		End if

                        Poliza = rs("poliza")

    			Sql3 = "Update Avisos Set "
   			Sql3 = Sql3 & " SubEntidad ='" & Siglas & TICKET & "',"
    			Sql3 = Sql3 & " SApertura=1"
    			Sql3 = Sql3 & " WHERE "
    			Sql3 = Sql3 & " ("
    			Sql3 = Sql3 & " Clinica = '" & Siglas & "'"
    			Sql3 = Sql3 & " OR"
   			Sql3 = Sql3 & " Poliza = '" & Poliza & "'"
    			Sql3 = Sql3 & " )"

               		cn.execute Sql3



                        txt_cedula = session("Cedula")
 
                        sql4 = "Insert Into TicketsAsociados ("
                        sql4 = Sql4 & " Entidad,"
			sql4 = Sql4 & " Subentidad,"
                        sql4 = Sql4 & " Ticket"
                        sql4 = Sql4 & " ) VALUES ("
                        sql4 = Sql4 & "'" & Txt_Cedula & "',"  
                        sql4 = Sql4 & "''," 
			Sql4 = Sql4 & "'" & Siglas & TICKET & "'"
                        Sql4 = Sql4 & " )"

                        cn.execute sql4   


                
                
             
		
	     		sql2 = " SELECT * " & _
               	               " FROM Personas " & _
               		       " WHERE Entidad = '" & session("Cedula") & "'" 
			  
				 Set rsx = Server.CreateObject("ADODB.Recordset")
         			 rsx.Open sql2, cn, 3, 3 
		 
	       
			 	 
		
		%>
</strong> 
<div id="Layer1" style="position:absolute; width:340px; height:24px; z-index:7; left: 9px; top: 1px;"> 
  <table width="101%" border="0">
    <tr> 
      <td><strong><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <%response.write "<STRONG>" & "<CENTER>" & "DATOS DEL ASEGURADO" & "</CENTER>" & "</STRONG>" 
	 response.write "<CENTER>" & "ASEGURADO Y POLIZA VIGENTE, CUMPLE TODOS LOS REQUISITOS" & "</CENTER>" 
	 response.write "<CENTER>" & "SERVICIO GARANTIZADO" & "</CENTER>" & "<br>"
	 response.write "<STRONG>" &"<CENTER>" & "TICKET: "&  SiGLAS  & TICKET & "</CENTER>" & "</STRONG>" & "<br>"%>
        </font></strong></td>
    </tr>
  </table>
</div>
<div id="Layer2" style="position:absolute; left:47px; top:84px; width:302px; height:106px; z-index:6"> 
  <table width="100%" border="0">
    <tr> 
      <td width="31%"><strong><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif">Nombre:</font></strong></td>
      <td width="69%"><strong><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <% =rsx.Fields("Nombre")%>
        </font></strong></td>
    </tr>
    <tr> 
      <td><strong><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif">Apellidos:</font></strong></td>
      <td><strong><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <%  =rsx.Fields("PrimerApellido") & " " & rsx.Fields("SegundoApellido")%>
        </font></strong></td>
    </tr>
    <tr> 
      <td><strong><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cedula:</font></strong></td>
      <td><strong><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <%  =rsx.Fields("Documento")%>
        </font></strong></td>
    </tr>
    <tr> 
      <td><strong><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif">FNacimiento:</font></strong></td>
      <td><strong><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <%  =rsx.Fields("FNacimiento")%>
        </font></strong></td>
    </tr>
    <tr> 
      <td><strong><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif">Sexo:</font></strong></td>
      <td><strong><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <%  =rsx.Fields("Sexo")%>
        </font></strong></td>
    </tr>
    <tr> 
      <td><strong><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif">Estado 
        Civil:</font></strong></td>
      <td><strong><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <% =rsx.Fields("EstadoCivil")%>
        </font></strong></td>
    </tr>
  </table>
</div>
<strong> 
<% end if

		rs.Close
		Set rs = Nothing
		
		cn.Close
		Set cn = Nothing

END IF
%>
</strong> 
<div id="Layer4" style="position:absolute; left:313px; top:213px; width:74px; height:28px; z-index:5"> 
  <strong><a href="sesion.asp"><img src="botonmenu.gif" width="57" height="28" border="0"></a> 
  </strong></div>
<p align="center">&nbsp;</p>
</BODY>
</HTML>
