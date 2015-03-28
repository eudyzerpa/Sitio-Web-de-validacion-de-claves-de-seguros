<%

   if session("LoggedIn") = 0 then
       response.redirect("login.asp?autorizado=falso")
    End if
 

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
		    response.Redirect("nologon.asp")
        else 
'			rs.MoveFirst 

'			Do While Not rs.EOF
'					response.write "USUARIO:"  & rs.Fields("Usuario")
'					response.write "CLAVE:" &  rs.Fields("Clave") & " " & _
'				rs.MoveNext
'			Loop
            session("Clinica")= rs.fields("Nombre")
			response.redirect("sesion.asp") 
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
<BODY>
<p>&nbsp;</p>
<p>&nbsp;</p>
<FORM METHOD="Post" name="Login" ACTION="Login.asp">
    <div align="center">
      <input type="hidden" name="consulta" value="true">
  </div>
    
  <H4 align="center" class="style1">&nbsp; </H4>
	<div align="center">
	  
    <TABLE width="279" BORDER=0>
      <TR>
        <TD class="style4 style5">CLAVE DE INGRESO
        <TD class="style4"><INPUT NAME="Usuario" SIZE="15"> 
      <TR>
        <TD COLSPAN=2 class="style4"><input name="Submit" type="Submit" value="Enviar">	
    </TABLE>
  </div>
</FORM>
</BODY>
</HTML>
