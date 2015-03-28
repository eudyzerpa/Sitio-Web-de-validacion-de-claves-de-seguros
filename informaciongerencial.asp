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
  <div align="center"><strong><font color="#000080" face="Verdana, Arial, Helvetica, sans-serif"> 
    <% response.Write(session("Clinica"))  %>
    </font></strong></div>
</div>

<div align="center">
  <div id="Layer1" style="position:absolute; width:565px; height:115px; z-index:3; left: 300px; top: 50px;"> 
    <table width="100%" border="0" align="center">
      <tr> 
        <td><p align="left"><font face="Verdana, Arial, Helvetica, sans-serif"><a href="casosabiertos.asp">Casos 
            Abiertos </a></font></p></td>
      </tr>
      <tr> 
        <td><font face="Verdana, Arial, Helvetica, sans-serif"><a href="casospresupuestados.asp">Casos 
          Presupuestados</a></font></td>
      </tr>
      <tr> 
        <td height="22"> <p align="left"><font face="Verdana, Arial, Helvetica, sans-serif"><a href="casosautorizados.asp">Casos 
            Autorizados </a></font></p></td>
      </tr>
    </table>
  </div>
</div>
<div id="Layer4" style="position:absolute; left:371px; top:129px; width:74px; height:28px; z-index:5"> 
  <a href="sesion.asp"><img src="botonmenu.gif" width="57" height="28" border="0"></a> 
</div>
</BODY>
</HTML>
