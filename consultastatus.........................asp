<HTML>
<HEAD><TITLE>Active Server Pages</TITLE>
<style type="text/css">
<!--
.style1 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 12;
	color: #000080;
}
.style4 {font-size: 12}
.style6 {font-family: Verdana, Arial, Helvetica, sans-serif; color: #000080; }
.style3 {font-size: 12px}
.style5 {font-family: Verdana, Arial, Helvetica, sans-serif; color: #000080; font-size: 12px; }
-->
</style>
</HEAD>
<BODY BGCOLOR=FFFFFF>
<FORM METHOD="Post" name="consultastatus" ACTION="consultastatus.asp">
    <div align="center">
      <p><input type="hidden" name="consulta2" value="true">
  </p></div>
    
  <center>
    <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif"> 
    <% response.Write(session("Clinica"))  %>
  </font></center>
  <H5 align="center" class="style1">AUTORIZACI&Oacute;N DE SERVICIO</H5>

	<div align="center">
       
    <div id="Layer1" style="position:absolute; width:245px; height:115px; z-index:1; left: 362px; top: 67px;"> 
      <font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;V 
      <input name="Tipo" type="radio" value="CIV" checked>
      E</font> <font size="2"> 
      <input type="radio" name="Tipo" value="CIE">
      </font> 
      <table width="100%" border=0>
        <tr> 
          <td width="70" class="style5">CEDULA&nbsp;
<td width="165" class="style1"><input name="Cedula" class="style2" size="14"> 
        <tr> 
          <td height="27" colspan=2 class="style1"><font size="2">CARTA AVAL</font> 
            <input name="carta" class="style2" size="14">
        <tr> 
          <td height="26" colspan=2 class="style4"><input name="Submit" type="submit" value="Enviar"> 
      </table>
      </div>
<H4 align="center" class="style1">&nbsp;</H4>
<div align="center"></div>
	  
  </div>
</FORM>
<%

if request.form("consulta2") = "true" then
        openstr = "driver={Microsoft Access Driver (*.mdb)};" & _
        "dbq=" & Server.MapPath("DBXSINRED.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr

  
        
        Entidad = request.form("Tipo") & request.form("Cedula")
                   
	
        sql = " SELECT * " & _
              " FROM Asegurados" & _
              " WHERE Entidad = '" & Entidad  & "'"
			  
		 
         Set rs = Server.CreateObject("ADODB.Recordset")
         rs.Open sql, cn, 3, 3 

         if rs.eof then
	 response.redirect("mensaje000090.asp")
         else 

			 	if rs.fields("Status") <> "" Then
			
					response.Redirect("mensaje000092.asp")
							
			 	else					
				    session("Cedula")= Entidad
				 	response.redirect("datosasegurado.asp")
			 	 end if
			
	 	end if

	 rs.Close
	 Set rs = Nothing
		
	 cn.Close
	 Set cn = Nothing

END IF
%>
</BODY>
</HTML>
