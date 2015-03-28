<%

if session("LoggedIn") = 0 then
       response.redirect("login.asp?autorizado=falso")
    End if

   if request.form("consulta") = "true" then
        openstr = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=rdmcds;Initial Catalog=DBXSINRED;Data Source=dell1600"
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr

        sql = " SELECT * " & _
              " FROM CasosAbiertos" & _
              " WHERE Ticket = '" & request.form("Usuario") & "'" 
			 
			  
	
        Set rs = Server.CreateObject("ADODB.Recordset")        		  
        rs.Open sql, cn, 3, 3


        if rs.eof then
		    response.Redirect("mensaje000011.asp")
                     Else 
        XDiagnostico = rs.fields("Diagnostico")
        session("ClavedeIngreso") = request.form("Usuario")

        rs.Close
	Set rs = Nothing
         
	
        if request.form("consulta") = "true" and XDiagnostico <> "" then
	 sqlupdate = " UPDATE casosabiertos " & _
                        " SET Estatus = 'POR AUTORIZAR'," & _
                        " FSolicitudDeAutorizacion = #" & Now & "#" & _
                        " WHERE Ticket = '" & session("ClavedeIngreso") & "'"
                     						
	 cn.Execute sqlupdate, raffected
         
         	if raffected > 0 then
             	 response.Redirect("mensaje000779.asp")

                 Dim tmpSiglas
  
    		tmpsiglas = session("ClavedeIngreso")


   		 Siglas = Mid(tmpsiglas, 1, 4) 'Extrae las Siglas
     
   		 Sql2 = "Update Avisos Set "
   		 Sql2 = Sql2 & " SubEntidad ='" & tmpSiglas & "',"
   		 Sql2 = Sql2 & " SSolicitudDeAutorizacion=1"
   		 Sql2 = Sql2 & " WHERE "
    		 Sql2 = Sql2 & " ("
   		 Sql2 = Sql2 & " Clinica = '" & Siglas & "'"
    		 Sql2 = Sql2 & " OR"
   		 Sql2 = Sql2 & " Poliza = '" & xPoliza & "'"
   		 Sql2 = Sql2 & " )"

               cn.execute Sql2
               
         	else
             	 response.Redirect("mensaje000070.asp")
         	END IF

             
       
         Else
         	response.Redirect("mensaje000778.asp") 	
           END IF 
   	END IF				
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
<p>&nbsp;</p>
<FORM METHOD="Post" name="ClaveIngreso" ACTION="solicitarautorizacion.asp">
    <div align="center">
      <input type="hidden" name="consulta" value="true">
  </div>
    
  <H4 align="center" class="style1">&nbsp; </H4>
	<div align="center">
	  
    <TABLE width="279" BORDER=0>
      <TR>
        <TD class="style4 style5">TICKET
        <TD class="style4"><INPUT NAME="Usuario" SIZE="15"> 
      <TR>
        <TD COLSPAN=2 class="style4"><input name="Submit" type="Submit" value="Enviar">	
    </TABLE>
  </div>
</FORM>
</BODY>
</HTML>

