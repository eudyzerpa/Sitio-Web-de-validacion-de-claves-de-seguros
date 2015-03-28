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
        "dbq=" & Server.MapPath("dbxsaps.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr
    
	  SQL = "UPDATE CasosAbiertos set MedicoTratante = '" & Session("Medico") & "'," & _
	                            " Diagnostico = '" & Session("Diagnostico") & "'" & _
	                            " WHERE CasosAbiertos.ClavedeIngreso = '" & session("ClavedeIngreso") & "'"
	                                
	                        
     response.Write(SQL)
     cn.Execute SQL
     Response.Write "Actualizacion exitosa"
	  
    ' if Affected > 0 then
    '     Response.Redirect("session.asp")
    ' else
    '     Response.Write "Ha ocurrido un error al actualizar el usuario, por favor intente de nuevo"
    ' end if

  %>


<p>&nbsp;</p>

<H4 align="center" class="style1">&nbsp;</h4>
</BODY>
</HTML>
