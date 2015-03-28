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

</HEAD>
<BODY Background="back2.jpg" vlink="black" link="black">

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
    Sql = Sql & " Personas.Documento,"
    Sql = Sql & " CasosAbiertos.MedicoTratante,"
    Sql = Sql & " CasosAbiertos.Presupuesto,"
    Sql = Sql & " CasosAbiertos.Diagnostico,"
    Sql = Sql & " CasosAbiertos.Poliza"
    
        
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
                        xPoliza  = rs.fields("Poliza")
		  			  		  	
		 End If
		 
		 Session("Medico")= request.Form("Txt_Medico")
		 Session("Diagnostico")= request.Form("Txt_Diagnostico")
		 
		 rs.close
		 Set rs = Nothing        

     if Request.Form("comportamiento") = "actualizar" then
                 xvariable = Request.Form("Txt_presupuesto")
			
         	 if xvariable = "" then
         	    xvariable = "0"
				else 
			    if isnumeric(xvariable) = false  then
					response.Redirect("mensaje000042.asp")
				end if	
          	 end if 
			    
				 
						  
         		sqlupdate = " UPDATE casosabiertos " & _
                      		    " SET Presupuesto = '" & xvariable & "'," & _
                      		    " Baremo = '" & Request.Form("Txt_cod_cargo") & "'," & _ 
                       		    " MedicoTratante = '" & Request.Form("Txt_Medico")  & "'," & _
                       		    " Estatus = 'POR AUTORIZAR'," & _
                                    " FPresupuesto = #" & Now & "#," & _
                       		    " Diagnostico = '" & Request.Form("Txt_Diagnostico") & "' " & _ 
                      		    " WHERE Casosabiertos.Ticket = '" & session("ClavedeIngreso") & "'"
         
         cn.Execute sqlupdate, raffected
         
         if raffected > 0 then
              response.Redirect("mensaje000093.asp")
         else
           response.Redirect("mensaje000034.asp")
         end if
     
      end if


    Dim tmpSiglas
  
    tmpsiglas = session("ClavedeIngreso")


    Siglas = Mid(tmpsiglas, 1, 4) 'Extrae las Siglas
     
    Sql2 = "Update Avisos Set "
    Sql2 = Sql2 & " SubEntidad ='" & tmpSiglas & "',"
    Sql2 = Sql2 & " SPresupuesto=1"
    Sql2 = Sql2 & " WHERE "
    Sql2 = Sql2 & " ("
    Sql2 = Sql2 & " Clinica = '" & Siglas & "'"
    Sql2 = Sql2 & " OR"
    Sql2 = Sql2 & " Poliza = '" & xPoliza & "'"
    Sql2 = Sql2 & " )"

    cn.execute Sql2
    




	          
  %>
<table width="100%" border="0">
  <tr> 
    <td width="11%"><strong><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif">Ticket:</font></strong></td>
    <td width="89%"> <font color="#000099" size="1"> 
      <% response.Write(session("ClavedeIngreso"))  %>
      </font></td>
  </tr>
  <tr> 
    <td><strong><font color="#000080" size="1" face="Verdana, Arial, Helvetica, sans-serif">Asegurado:</font></strong></td>
    <td><font color="#000080" size="1"> 
      <% = xNombre & " " & xapellido %>
      </font></td>
  </tr>
  <tr> 
    <td><strong><font color="#000080" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cedula:</font></strong></td>
    <td><font color="#000080" size="1"> 
      <% = xcedula %>
      </font></td>
  </tr>
</table>

<div id="Layer1" style="position:absolute; width:450px; height:115px; z-index:1; left: 8px; top: 94px;"> 
  <form name="Actualizar" method="post" action="actualizardiagnosticodelcaso.asp">
    <input type="hidden" name="comportamiento" value="actualizar">
    <table width="55%" border="0" >
      <tr> 
        <td width="36%" height="24"><strong><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif">M&eacute;dico:</font></strong></td>
        <td width="64%"> <font color="#000099" size="1"> 
          <input type="text" size="15" name="Txt_Medico">
          </font></td>
      </tr>
      <tr> 
        <td width="36%" height="24"><strong><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif">Baremo:</font></strong></td>
        <td width="64%"> <font color="#000099" size="1"> 
          <input type="text" size="15" name="Txt_cod_cargo">
          </font></td>
      </tr>

      <tr> 
        <td height="24"><strong><font color="#000080" size="1" face="Verdana, Arial, Helvetica, sans-serif">Diagnostico:</font></strong></td>
        <td> <font color="#000080" size="1"> 
          <input type="text" size="15" name="Txt_Diagnostico">
          </font></td>
      </tr>
      <tr> 
        <td height="24"><strong><font color="#000080" size="1" face="Verdana, Arial, Helvetica, sans-serif">Presuspuesto:</font></strong></td>
        <td> <font color="#000080" size="1"> 
          <input type="text" size="15" name="Txt_presupuesto">
          </font></td>
      </tr>
      <tr> 
        <td height="25"><font color="#000080" size="1"><a href="JavaScript:document.Actualizar.submit();"><img src="enviar.gif" width="57" height="28" border="0"></a></font></td>
        <td><font color="#000080" size="1">&nbsp;</font></td>
      </tr>
    </table>
    <p>&nbsp;&nbsp;&nbsp;&nbsp; </p>
    </form>
</div>
<H4 align="center" class="style1">&nbsp;</h4>
</BODY>
</HTML>
