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
	
  <H4 align="center" class="style1"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">La 
    c�dula fue encontrada en el registro de asegurados </font></H4>
  <H4 align="center" class="style1"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Para 
    servicio debe llamar a la empresa aseguradora</font></H4>
</FORM>

<div id="Layer4" style="position:absolute; left:462px; top:101px; width:74px; height:28px; z-index:5"> 
  <a href="sesion.asp"><img src="botonmenu.gif" width="57" height="28" border="0"></a> 
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
