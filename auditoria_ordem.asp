<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!-- #include file="config.asp" -->
<!--#include file="styles.asp"-->
<% If Session("LoggedIn") = True Then
	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from os order by os_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic

	ordem=CInt(request.querystring("pesquisa_os")) %>
<body><center>
<table border="1" width="780" height="102">
    <tr>
      <td class="fundo1" width="100" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="680" height="102" align="center">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Controle de Ordem de Serviço</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Área Restrita aos Administradores - Auditoria</b></font>
<script language="JavaScript">
<!--
var months=new Array(13);
months[1]="Janeiro";
months[2]="Fevereiro";
months[3]="Mar&ccedil;o";
months[4]="Abril";
months[5]="Maio";
months[6]="Junho";
months[7]="Julho";
months[8]="Agosto";
months[9]="Setembro";
months[10]="Outubro";
months[11]="Novembro";
months[12]="Dezembro";
var time=new Date();
var lmonth=months[time.getMonth() + 1];
var date=time.getDate();
year=time.getFullYear();
var today = new Date();
var hrs = today.getHours();
document.write("<leftt>");
document.write("<B>");
document.write("<font face=Verdana size=1 color=#000000>");
if (hrs < 6)
document.write("Bom Dia!!! -");
else if (hrs < 12)
document.write("Bom Dia!!! -");
else if (hrs < 18)
document.write("Boa Tarde!!! -");
else
document.write("Boa Noite!!! -");   
document.write(" Porto Velho, ");
document.write(date + "  de ");
document.write(lmonth + " de  " + year + "</left></B>");
//-->
	  </script>	  
      </b></font></p></td>
    </tr>
</table>
<table width="780" align="center" border="1" height="23">
<tr>
	  <td class="fundo1" width="40" height="23" align="center"><b><font size="2">Número</td>
      <td class="fundo1" width="100" height="23" align="center"><b><font size="2">Solicitante</td>
      <td class="fundo1" width="100" height="23" align="center"><b><font size="2">Esquadrão</td>
      <td class="fundo1" width="100" height="23" align="center"><b><font size="2">Seção</td>
      <td class="fundo1" width="100" height="23" align="center"><b><font size="2">IP</td>
      <td class="fundo1" width="100" height="23" align="center"><b><font size="2">Host</td>
      <td class="fundo1" width="100" height="23" align="center"><b><font size="2">Login</td>
	  <td class="fundo1" width="100" height="23" align="center"><b><font size="2">Server</td>
	  <td class="fundo1" width="40" height="23" align="center"><b><font size="2">N° OS</td>
</tr>
<% 	HowMany = 0
	Do While Not rsbanco.EOF And HowMany < rsbanco.PageSize
		If rsbanco("os_numero") = ordem then%>
<tr>
	  <td class="fundo5" width="40" height="27" align="center"><a href="solicitacao.asp?codsolic=<% = rsbanco("os_codigo") %>"><font color="#7F0D11" size="2"><b><% = rsbanco("os_codigo") %></a></td>
      <td class="fundo5" width="100" height="27" align="center"><font size="2"><% = rsbanco("os_solicmilitar") %></td>
      <td class="fundo5" width="100" height="27" align="center"><font size="2"><% = rsbanco("os_solicesquadrao") %></td>
      <td class="fundo5" width="100" height="27" align="center"><font size="2"><% = rsbanco("os_solicsecao") %></td>
      <td class="fundo5" width="100" height="27" align="center"><font size="2"><% = rsbanco("os_ip") %></td>
      <td class="fundo5" width="100" height="27" align="center"><font size="2"><% = rsbanco("os_host") %></td>
      <td class="fundo5" width="100" height="27" align="center"><font size="2"><% = rsbanco("os_logon") %></td>
	  <td class="fundo5" width="100" height="27" align="center"><font size="2"><% = rsbanco("os_server") %></td>      
      <td class="fundo5" width="40" height="29" align="center"><% If rsbanco("os_numero") > 0 Then %><a href="os.asp?codsolic=<% = rsbanco("os_codigo") %>"><% End If %><font color="#7F0D11" size="2"><b><% = rsbanco("os_numero") %></td>	
</tr>
<% 		HowMany = HowMany + 1
		rsbanco.MoveNext
	Else
		rsbanco.MoveNext
	End If
	Loop %>
</table>
<table border="1" width="780" height="30">
    <tr>
      	<form method="GET" action="auditoria_solicitacao.asp">
      		<td class="fundo1" width="150" height="30" align="center"><b><font size="2">N° da Solicitação</td>
      		<td class="fundo3" width="120" height="30" align="center"><input type="text" name="pesquisa_sol" size="10" style="border-style: inset; border-width: 5; text-align:center"></td>
      		<td class="fundo5" width="120" height="30" align="center"><input type="submit" value="Procurar"></td>
      	</form>
      	<form method="GET" action="auditoria_ordem.asp">
      		<td class="fundo1" width="150" height="30" align="center"><b><font size="2">N° Ordem de Serviço</td>
      		<td class="fundo3" width="120" height="30" align="center"><input type="text" name="pesquisa_os" size="10" style="border-style: inset; border-width: 5; background-color:#00FFFF; text-align:center"></td>
      		<td class="fundo5" width="120" height="30" align="center"><input type="submit" value="Procurar"></td>
      	</form>
    </tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="780" align="center">		  
	<tr>
		<form action="auditoria.asp"><td align="center" width="780"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;"></td></form>
	</tr>
</table>
<!--#include file="rodape.asp"--></center></body></html>
<% Else
	Response.Redirect("admin.asp")
End If %>