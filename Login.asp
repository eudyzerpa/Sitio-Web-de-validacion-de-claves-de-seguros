<%

    if request.form("consulta") = "true" then
        openstr = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=rdmcds;Initial Catalog=DBXSINRED;Data Source=dell1600"        
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr


        sql = " SELECT * " & _
              " FROM Usuarios" & _
              " WHERE Usuario= '" & request.form("Usuario") & _
	      "' AND Clave ='" & request.form("Clave") & "'" 
			  
			  
           '			  response.Write "SQL--------->" & sql & "<---------" 
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open sql, cn, 3, 3 

        if rs.eof  then
		    response.Redirect("mensaje000080.asp")
        else 
                        
                        session("Clinica")= rs.fields("AsociadoA")
                        Session("Usuario")= request.form("Usuario") 
                        session("LoggedIn") = 1
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

    <% if request.querystring("autorizado") = "falso" then
          response.redirect("mensaje0000100.asp")
       End if
     %> 

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
<BODY Background="back.jpg">
<div align="left"><strong><font color="#000099" size="9" face="Lucida Sans">SIN</font><font color="#000066" size="9" face="Lucida Sans"><font color="#FF0000">RED</font></font></strong> 
</div>

<FORM METHOD="Post" name="Login" ACTION="Login.asp">
    
  <div align="center"> 
    <input type="hidden" name="consulta" value="true">
    <font color="#FF2000" size="8" face="Lucida Sans"></font></div>
    
  <div align="center">
	  
    <TABLE height="91" BORDER=0>
      <TR>
        <TD class="style4 style5"><div align="right"><span class="style3"><strong><font color="#FFFFFF">USUARIO</font></strong></span><font color="#FFFFFF"><strong>:</strong></font><strong> 
            </strong> </div>
        <TD class="style4"><INPUT NAME="Usuario" SIZE="15">
	  <TR>
        <TD class="style4"><div align="right"><span class="style6"><strong><font color="#FFFFFF">CLAVE:</font></strong> 
            </span> </div>
        <TD class="style4"><INPUT TYPE="Password" NAME="Clave" SIZE="15">
	  <TR><TD COLSPAN=2 class="style4"><input name="Submit" type="Submit" value="Enviar">	    
	    </TABLE>
  </div>
</FORM>
</BODY>
</HTML>
