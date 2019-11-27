<html>
<head>
<title>.:.:.:. COMPOS WEB .:.:.:.</title>
</head>
<body>
<!-- #include file="config.asp" -->
<br><center><b>Exclusão de Dados</b></center>
<br><br>
<table size ="80%" align="center" border="1">
<tr>
<%
	for each campo in rsbanco.fields
%>
	<td><b><% = campo.name %></b>&nbsp;</td>
<%
	next
%>
</tr>
<% 
	While not rsbanco.eof
%>
    <tr>
	  <td width="80" height="29" align="center"><% = rsbanco("cod_sol") %></td>
      <td width="80" height="29"><% = rsbanco("esquadrao") %></td>
      <td width="80" height="29"><% = rsbanco("secao") %></td>
      <td width="80" height="29" align="center"><% = rsbanco("periferico") %></td>
      <td width="280" height="29"><% = rsbanco("descricao") %></td>
      <td width="100" height="29"><% = rsbanco("solicitante") %></td>
    </tr>
<%
		rsbanco.movenext
	Wend
	
		set banco = nothing
		set rsbanco = nothing
%>
</table><center>
<form action="del.asp" method="POST"><font size="2" color="#000000"><b>Entre com o Código para a exclusão:&nbsp;&nbsp;&nbsp;
	<input type="text" name="exclui" size="6">&nbsp;&nbsp;&nbsp;
	<input type="submit" value="Excluir Registro" name="B1"></b></font>
</form></center>
<br>
<center><a href="index.htm">Volta à Página Principal</a></center>
<div align="center">
  <table border="0" cellpadding="0" cellspacing="0" width="780" height="40">
    <tr>
      	<TD WIDTH=780 HEIGHT=40 COLSPAN=7 align="center" style="color:606060" class="tah10">
Copyright (c) 2006. Hallyz Cia & Ltda. Todos os direitos reservados.</TD>
    </tr>
  </table>
</div>
</body>
</html>