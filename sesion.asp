<%

    if session("LoggedIn") = 0 then
       response.redirect("login.asp?autorizado=falso")
    End if

  

    if request.form("consulta") = "true" then
        openstr = "driver={Microsoft Access Driver (*.mdb)};" & _
        "dbq=" & Server.MapPath("DBXSINRED.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr

        sql = " SELECT * " & _
              " FROM Usuarios" & _
              " WHERE Usuario= '" & request.form("Usuario") & _
			  "' AND Clave ='" & request.form("Clave") & "'" 
			  
           '			  response.Write "SQL--------->" & sql & "<---------" 
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open sql, cn, 3, 3 

        if rs.eof AND rs.BOF then
		    response.Redirect("nologon.asp")
        else 
'			rs.MoveFirst 

'			Do While Not rs.EOF
'					response.write "USUARIO:"  & rs.Fields("Usuario")
'					response.write "CLAVE:" &  rs.Fields("Clave") & " " & _
'				rs.MoveNext
'			Loop
			response.redirect("consultastatus.asp") 
	end if

		rs.Close
		Set rs = Nothing
		
		cn.Close
		Set cn = Nothing

     END IF
%>
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
<div id="Layer2" style="position:absolute; left:133px; top:27px; width:553px; height:20px; z-index:2"> 
  <div align="center"><font color="#000000"><strong><font face="Verdana, Arial, Helvetica, sans-serif"> 
    <% response.Write(session("Clinica"))  %>
    </font></strong></font></div>
</div>
<div align="center"> 
  <div id="Layer1" style="position:absolute; width:565px; height:115px; z-index:3; left: 300px; top: 50px;"> 
    <table width="100%" border="0" align="center">
      <tr> 
        <td><p align="left"><font color="#000000"><strong><font face="Verdana, Arial, Helvetica, sans-serif"><a href="consultastatus.asp">Abrir 
            casos</a></font></strong></font></p></td>
      </tr>
      <tr> 
        <td><font color="#000000"><strong><font face="Verdana, Arial, Helvetica, sans-serif"><a href="clavedeingresodiagnostico.asp">Diagnostico 
          del caso</a></font></strong></font></td>
      </tr>
      <tr> 
        <td><font color="#000000"><strong><font face="Verdana, Arial, Helvetica, sans-serif"><a href="solicitarautorizacion.asp">Solicitud  
         de autorización</a></font></strong></font></td>
      </tr>
      <tr> 
        <td height="22"> <p align="left"><font color="#000000"><strong><font face="Verdana, Arial, Helvetica, sans-serif"><a href="casosclinica.htm">Mostrar 
            casos cl&iacute;nica actual</a></font></strong></font></p></td>
      </tr>
      <tr> 
        <td height="21"> <p align="left"><font color="#000000"><strong><font face="Verdana, Arial, Helvetica, sans-serif"><a href="clavedeingresoclinicaselectiva.asp">Mostrar 
            casos cl&iacute;nica otra</a></font></strong></font></p></td>
      </tr>
      <tr> 
        <td height="21"><font color="#000000"><strong><font face="Verdana, Arial, Helvetica, sans-serif"><a href="clavedeingresomedicotratante.asp">Mostrar 
          casos m&eacute;dico</a></font></strong></font></td>
      </tr>
      <tr> 
        <td height="21"><font color="#000000"><strong><font face="Verdana, Arial, Helvetica, sans-serif"><a href="clavedeingresomostrarcargos.asp">Mostrar 
          cargos</a></font></strong></font></td>
      </tr>
      <tr> 
        <td height="22"><font color="#000000"><strong><font face="Verdana, Arial, Helvetica, sans-serif"><a href="clavedeingresonoautorizarcasos.asp">No 
          autorizar casos</a></font></strong></font></td>
      </tr>
      <tr> 
        <td height="21"><font color="#000000"><strong><font face="Verdana, Arial, Helvetica, sans-serif"><a href="Informaciongerencial.asp">Informaci&oacute;n 
          Gerencial</a></font></strong></font></td>
      </tr>
      <tr> 
        <td height="21"><font color="#000000"><strong><font face="Verdana, Arial, Helvetica, sans-serif"><a href="clavedeingresocerrarcasos.asp">Solicitud 
          de Liquidación de Casos</a></font></strong></font></td>
      </tr>
    </table>
  </div>
</div>

</BODY>
</HTML>
