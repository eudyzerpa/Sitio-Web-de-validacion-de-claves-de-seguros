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
        openstr = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=rdmcds;Initial Catalog=DBXSINRED;Data Source=dell1600"
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr
    

       Sql = ""
       Sql = Sql & " Select "
       Sql = Sql & " CasosAbiertos.Clinica,"
       Sql = Sql & " COUNT(*) as Total,"
       Sql = Sql & " SUM(MontoAutorizado) as Monto"
       Sql = Sql & " From "
       Sql = Sql & " CasosAbiertos"
       Sql = Sql & " WHERE CasosAbiertos.Autorizado = 'SI'"
       Sql = Sql & " GROUP BY CasosAbiertos.Clinica"
       
'sql = " SELECT CasosAbiertos.Clinica, COUNT(*) as Total FROM CasosAbiertos WHERE CasosAbiertos.Fpresupuesto <> NULL GROUP BY CasosAbiertos.Clinica" 
	
			  
		 
         Set rs = Server.CreateObject("ADODB.Recordset")
         rs.Open sql, cn, 3, 3 



 
%>

<p>&nbsp;</p><table width="100%" border="1">
  <tr bgcolor="#009933" class="style1" > 
      
    <th width="61" height="27"><div align="center"><font color="#FFFFFF">Clinica</font></div></th>
    <th width="49" height="27"> <div align="center" class="style9"><font color="#FFFFFF">Autorizados</font></div></th> 
    <th width="61" height="27"><div align="center"><font color="#FFFFFF">Monto</font></div></th>
     
    </tr>
    <% if Not rs.EOF then %>
    <% Do while Not rs.EOF %>
    <tr class="style1"> 
      <td><div align="left"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%= rs("Clinica") %></font></div></td>
      <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%= rs("Total") %></font></div></td>
      <td><div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%= rs("Monto") %></font></div></td>
    </tr>
    
     
   
    <% rs.MoveNext
	   Loop
	 %>
    <% Else 
	    response.redirect("mensaje000036.asp")
	 end if
	 %>
  </table>

<H4 align="center"><font color="#000080" face="Verdana, Arial, Helvetica, sans-serif"><!-- <% response.Write(session("Clinica")) %> --></font></h4>
<div id="Layer1" "style="position:absolute; width:845px; height:98px; z-index:1; left: 9px; top: 143px; background-color: #000000; layer-background-color: #000000; border: 1px none #000000; visibility: visible; overflow: scroll;"> 
</div>
<div id="Layer2" style="position:absolute; left:182px; top:21px; width:495px; height:42px; z-index:2"> 
  <div align="center"><font color="#400080" face="Verdana, Arial, Helvetica, sans-serif"><strong>CASOS 
    AUTORIZADOS</strong></font></div>
</div>
</BODY>
</HTML>
