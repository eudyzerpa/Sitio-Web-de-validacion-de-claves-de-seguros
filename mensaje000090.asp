<%

    if request.form("consulta") = "true" then
        openstr = "driver={Microsoft Access Driver (*.mdb)};" & _
        "dbq=" & Server.MapPath("dbxsaps.mdb")
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
		    response.write "La contrase�a es incorrecta, verifique haber escrito correctamente el Nombre de Usuario y la Clave."
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
<center>
  <div id="Layer2" style="position:absolute; width:200px; height:115px; z-index:2; left: 397px; top: 152px;"> 
    <h4><font color=#FF0000 size="2" face="Verdana, Arial, Helvetica, sans-serif"> (90) La C&eacute;dula no se encuentra en ninguna 
      P&oacute;liza de nuestros asegurados.</FONT></h4>
    <p><font color="#000066" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
      </strong> </font> </p>
</div>
  <div id="Layer4" style="position:absolute; left:452px; top:229px; width:74px; height:28px; z-index:5"> 
    <a href="sesion.asp"><img src="botonmenu.gif" width="57" height="28" border="0"></a> 
  </div>
</center>
</BODY>
</HTML>