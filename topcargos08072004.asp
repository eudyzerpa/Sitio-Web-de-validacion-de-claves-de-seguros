<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
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

<body>
<div id="Layer1" style="position:absolute; left:142px; top:12px; width:547px; height:22px; z-index:1"> 
  <table width="118%" border="0" align="center">
    <tr> 
      <td width="54%"><div align="center"><font color="#000080" size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
          <% response.Write(session("Clinica")& "<br>" & "Clave:" & " " & session("ClavedeIngreso")) %>
          </font></div></td>
    </tr>
  </table>
</div>
<center>
  <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif"> </font> 
</center>
</body>
</html>
