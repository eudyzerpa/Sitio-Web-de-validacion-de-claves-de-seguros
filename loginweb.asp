<html>
<head>
<title>ultra4</title>
<style type="text/css">
<!--
	body,td		{ font-family: arial, helvetica; font-size: 11px; color: #444444; }
	input.text	{ font-size: 10px; }
	.bottom		{ font-size: 10px; color: #c0c0c0; }
	a:link		{ color: #F26250; text-decoration: none; }
	a:hover		{ color: #F26250; text-decoration: underline; }
	a:visited	{ color: #F26250; text-decoration: none; }
	form		{ margin-top: 0; margin-bottom: 0; }
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
</head>
<body rightmargin="0" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" rightmargin="0" background="i/bg2.jpg">

<div id="Layer1" style="position:absolute; left:11px; top:11px; width:144px; height:122px; z-index:1"><img src="i/GloboCiberDyneGIF.gif" width="118" height="118"></div>
<div id="Layer2" style="position:absolute; left:281px; top:-71px; width:428px; height:89px; z-index:2"><img src="i/sinred.gif" width="454" height="340"></div>
<div id="Layer3" style="position:absolute; left:586px; top:241px; width:348px; height:243px; z-index:3"> 
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
		    response.Redirect("mensaje000080.asp")
        else 
'			rs.MoveFirst 

'			Do While Not rs.EOF
'					response.write "USUARIO:"  & rs.Fields("Usuario")
'					response.write "CLAVE:" &  rs.Fields("Clave") & " " & _
'				rs.MoveNext
'			Loop
            session("Clinica")= rs.fields("AsociadoA")
			response.redirect("sesion.asp") 
	end if
        
		rs.Close
		Set rs = Nothing
		
		cn.Close
		Set cn = Nothing

     END IF
%>
  <FORM METHOD="Post" name="Login" ACTION="login.asp">
    <div align="center">
      <input type="hidden" name="consulta" value="true">
  </div>
	  <TABLE BORDER=0>
	  <TR><TD class="style4 style5"><span class="style3">USUARIO</span>: 
	  <TD class="style4"><INPUT NAME="Usuario" SIZE="15">
	  <TR><TD class="style4"><span class="style6">CLAVE: </span>
	  <TD class="style4"><INPUT TYPE="Password" NAME="Clave" SIZE="15">
	  <TR><TD COLSPAN=2 class="style4"><input name="Submit" type="Submit" value="Enviar">	    
	    </TABLE>
  </div>
</FORM>

</div>
<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
	<td width="100%">
		<table width="100%" cellpadding="0" cellspacing="0" border="0">
		<tr>
			<td><img src="i/top1.jpg"></td>
			<td background="i/top1bg.jpg" width="100%"><img src="i/spacer.gif"></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td background="i/top2bg.jpg">
		<table width="100%" cellpadding="0" cellspacing="0" border="0">
		<tr>
			
          <td><img src="i/spacer.gif" width="30" height="8"></td>
			<td><a href="#"><img src="i/l1.jpg" border="0"></a></td>
			<td><img src="i/top2div.jpg"></td>
			<td><a href="#"><img src="i/l2.jpg" border="0"></a></td>
			<td><img src="i/top2div.jpg"></td>
			<td><a href="#"><img src="i/l3.jpg" border="0"></a></td>
			<td><img src="i/top2div.jpg"></td>
			<td><a href="#"><img src="i/l4.jpg" border="0"></a></td>
			<td><img src="i/top2div.jpg"></td>
			<td><a href="#"><img src="i/l5.jpg" border="0"></a></td>
			<td width="100%"><img src="i/spacer.gif"></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr> 
          <!--

	The image 'top3.jpg' contains the picture of the handheld. If you want
	this image without the handheld, change 'top3.jpg' to 'top3-blank.jpg'
	in the line immediately following this comment.

-->
          <td><img src="i/top3.jpg"></td>
          <td><img src="i/top3div.jpg"></td>
        </tr>
      </table>
	</td>
</tr>
<tr>
	<td background="i/top4bg.jpg"><img src="i/top4bg.jpg"></td>
</tr>
<tr>
	<td background="i/bg1.jpg">
		<table width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr> 
          <td><img src="i/spacer.gif" width="6" height="1"></td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td width="100%"><img src="i/spacer.gif"></td>
        </tr>
        <tr> 
          <td height="35"><img src="i/spacer.gif" width="1" height="13"></td>
        </tr>
      </table>
	</td>
</tr>
<tr>
	<td background="i/top5bg.jpg"><img src="i/top5bg.jpg"></td>
</tr>
<tr>
	<td>
		<table width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr> 
          <td rowspan="3"><img src="i/spacer.gif" width="28"></td>
          <td><img src="i/spacer.gif" height="1" width="1"></td>
        </tr>
        <tr> 
          <td width="100%" valign="top" class="bottom">&nbsp;</td>
        </tr>
        <tr> 
          <td><img src="i/spacer.gif" height="15" width="1"></td>
        </tr>
      </table>
	</td>
</tr>
</table>
</body>
</html>