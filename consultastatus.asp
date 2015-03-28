<%

   if session("loggedIn") = 0 then
       response.redirect("login.asp?autorizado=falso")
    end if
%>

<HTML>
<HEAD><TITLE>Sistema INterconectado de Envío y Recepción de Datos</TITLE>
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
</HEAD>
<BODY Background="back2.jpg" vlink="black" link="black">
<FORM METHOD="Post" name="consultastatus" ACTION="consultastatus.asp">
  <div align="center"> 
    <p><strong> 
      <input type="hidden" name="consulta2" value="true">
      </strong><strong><font color="#000080" face="Verdana, Arial, Helvetica, sans-serif"> 
      <% response.Write(session("Clinica"))  %>
      </font></strong> </p>
  </div>
  <!--  <H5 align="center" class="style1"><strong>AUTORIZACI&Oacute;N DE SERVICIO</strong></H5>-->
  <div align="center"> 
    <div id="Layer1" style="position:absolute; width:296px; height:97px; z-index:1; left: 360px; top: -19px;"> 
      <strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;V 
      <input name="Tipo" type="radio" value="CIV" checked>
      E</font> <font size="2"> 
      <input type="radio" name="Tipo" value="CIE">
      </font> </strong> 
      <table width="103%" border=0>
        <tr> 
          <td width="89" height="24" class="style5"><strong><font size="2">CEDULA</font>&nbsp; 
            </strong> 
          <td width="206" class="style1"><strong> 
            <input name="Cedula" class="style2" size="12">
            </strong> 
        <tr> 
          <td height="24" colspan=2 class="style1"><strong><font size="2">CARTA 
            AVAL</font> 
            <input name="carta" class="style2" size="12">
            </strong> 
        <tr> 
          <td height="26" colspan=2 class="style4"><strong> 
            <input name="Submit" type="submit" value="Enviar">
            </strong> 
      </table>
    </div>
    <H4 align="center" class="style1">&nbsp;</H4>
    <div align="center"></div>
  </div>
</FORM>
<strong> 
<%

Dim xlongitud,xvariable,entidad


if request.form("consulta2") = "true" then
        openstr = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=rdmcds;Initial Catalog=DBXSINRED;Data Source=dell1600"        
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr
        
        Entidad =  request.form("Cedula")
        xlongitud = LEN(Entidad)
                 
			Select Case xlongitud
        
				Case 0 
					Response.Redirect("mensaje000090.asp")
				Case 1
					xvariable ="00000000" & Entidad
				Case 2
					xvariable ="0000000" & Entidad
				Case 3
					xvariable ="000000" & Entidad
				Case 4
					xvariable ="00000" & Entidad
				Case 5
					xvariable ="0000" & Entidad
				Case 6
					xvariable ="000" & Entidad
				Case 7
					xvariable ="00" & Entidad
				Case 8
					xvariable ="0" & Entidad
				Case 9
					xvariable = Entidad
			End Select
        
        xvariable = request.form("Tipo") & xvariable
       
                 
	
        sql = " SELECT * " & _
              " FROM Asegurados" & _
              " WHERE Entidad = '" & xvariable & "'"
			  
		 
         Set rs = Server.CreateObject("ADODB.Recordset")
         rs.Open sql, cn, 3, 3 

         if rs.eof then
	 response.redirect("mensaje000090.asp")
         else 

			 	if rs.fields("Status") <> "" Then
			
					response.Redirect("mensaje000092.asp")
							
			 	else					
				    session("Cedula")= xvariable
				 	response.redirect("datosasegurado.asp")
			 	 end if
			
	 	  end if

	 rs.Close
	 Set rs = Nothing
		
	 cn.Close
	 Set cn = Nothing

END IF
%>
</strong> 
</BODY>
</HTML>
