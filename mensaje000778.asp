<% IF Request.Form = "" THEN %>

<HTML>
<HEAD><TITLE>Active Server Pages</TITLE>
<style type="text/css">
<!--
.style1 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	color: #000080;
}
-->
</style>
</HEAD>
<BODY Background="back2.jpg" vlink="black" link="black">
<CENTER>
<FORM METHOD=Post ACTION=file:///Y|/Ejemplo5a.asp>
	
  <H4 align="center" class="style1"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Debe existir un diagnostico
    </font></H4>
  </FORM>

<div id="Layer4" style="position:absolute; left:441px; top:84px; width:104px; height:28px; z-index:5"> 
  <table width="83%" border="0">
    <tr> 
      <td width="27%"><div align="center"><a href="sesion.asp"><img src="botonmenucargostarifados.gif" width="57" height="28" border="0"></a></div></td>
      <td width="73%"><div align="center"><a href="clavedeingresodiagnostico.asp"><img src="botonvolver.gif" width="57" height="28" border="0"></a></div></td>
    </tr>
  </table>
</div>
</BODY>
</HTML>

<% ELSE

	IF (Request.Form ("Usuario") = "Luis" AND Request.Form ("Clave") = "31416") _
	OR (Request.Form ("Usuario") = "Ale" AND Request.Form ("Clave") = "Luckan") THEN
		Session ("Autentificado") = True
		Response.Cookies ("Usuario") = Request.Form ("Usuario")
		Response.Redirect "Ejemplo5b.asp"
	ELSE
		Response.Redirect "Ejemplo5a.asp"
	END IF

END IF %>
