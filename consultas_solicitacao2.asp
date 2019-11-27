<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!-- #include file="config.asp" -->
<!--#include file="styles.asp"-->
<% If Session("LoggedIn") = True Then
	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from os order by os_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic

	solicitacao=CInt(request.querystring("pesquisa_sol")) %>
<body><center>
<table border="1" width="750" height="102">
    <tr>
      <td class="fundo1" width="100" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="650" height="102" align="center">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Controle de Ordem de Serviço</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Consulta Completa das Solicitações</b></font>
	  <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
	  <p style="margin-top: 0; margin-bottom: 0"><font color="#000000"><b>
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
<!--#include file="statusos.asp"-->
<table width="750" align="center" border="1" height="82">
<tr>
	  <td class="fundo1" width="50" height="23" align="center"><b><font size="2">Número</td>
      <td class="fundo1" width="120" height="23" align="center"><b><font size="2">Data</td>
      <td class="fundo1" width="90" height="23" align="center"><b><font size="2">Periférico</td>
      <td class="fundo1" width="140" height="23" align="center"><b><font size="2">Solicitante</td>
      <td class="fundo1" width="70" height="23" align="center"><b><font size="2">Esquadrão</td>
      <td class="fundo1" width="140" height="23" align="center"><b><font size="2">Seção</td>
      <td class="fundo1" width="50" height="23" align="center"><b><font size="2">Ramal</td>
	  <td class="fundo1" width="40" height="23" align="center"><b><font size="2">Status</td>
	  <td class="fundo1" width="50" height="23" align="center"><b><font size="2">N° OS</td>	        
</tr>
<% 	HowMany = 0
	Do While Not rsbanco.EOF And HowMany < rsbanco.PageSize
		If rsbanco("os_codigo") = solicitacao then%>
<tr>
	  <td width="50" height="27" align="center" bgcolor="#FFFF00"><a href="solicitacao.asp?codsolic=<% = rsbanco("os_codigo") %>"><font color="#7F0D11" size="2"><b><% = rsbanco("os_codigo") %></a></td>
      <td class="fundo5" width="120" height="27" align="center"><font size="2"><% = rsbanco("os_solicdata") %></td>
      <td class="fundo5" width="90" height="27" align="center"><font size="2"><% = rsbanco("os_solicperiferico") %></td>
      <td class="fundo5" width="140" height="27" align="center"><font size="2"><% = rsbanco("os_solicmilitar") %></td>
      <td class="fundo5" width="70" height="27" align="center"><font size="2"><% = rsbanco("os_solicesquadrao") %></td>
      <td class="fundo5" width="140" height="27" align="center"><font size="2"><% = rsbanco("os_solicsecao") %></td>
      <td class="fundo5" width="50" height="27" align="center"><font size="2"><% = rsbanco("os_solicramal") %></td>
<% If rsbanco("os_status") = 1 then %>
	  <td class="fundo5" width="40" height="27" align="center"><img border="0" src="bolaverde.gif"></td>      
<% Else
		If rsbanco("os_status") = 2 then %>
      <td class="fundo5" width="40" height="27" align="center"><img border="0" src="bolaamarela.gif"></td>
		<% Else
			If rsbanco("os_status") = 3 then %>
      <td class="fundo5" width="40" height="27" align="center"><img border="0" src="bolaazul.gif"></td>		
			<% Else %>
      <td class="fundo5" width="40" height="27" align="center"><img border="0" src="bolavermelha.gif"></td>
			<% End If
		End If
	End If %>
      <td class="fundo5" width="50" height="29" align="center"><% If rsbanco("os_numero") > 0 Then %><a href="os.asp?codsolic=<% = rsbanco("os_codigo") %>"><% End If %><font color="#7F0D11" size="2"><b><% = rsbanco("os_numero") %></td>	
</tr>
<% 		HowMany = HowMany + 1
		rsbanco.MoveNext
	Else
		rsbanco.MoveNext
	End If
	Loop %>
</table>
<table border="1" width="750" height="30">
    <tr>
      	<form method="GET" action="consultas_solicitacao2.asp">
      		<td class="fundo1" width="140" height="30" align="center"><b><font size="2">N° da Solicitação</td>
      		<td class="fundo3" width="125" height="30" align="center"><input type="text" name="pesquisa_sol" size="10" style="border-style: inset; border-width: 5; background-color:#FFFF00; text-align:center"></td>
      		<td class="fundo5" width="110" height="30" align="center"><input type="submit" value="Procurar"></td>
      	</form>
      	<form method="GET" action="consultas_ordem2.asp">
      		<td class="fundo1" width="140" height="30" align="center"><b><font size="2">N° Ordem de Serviço</td>
      		<td class="fundo3" width="125" height="30" align="center"><input type="text" name="pesquisa_os" size="10" style="border-style: inset; border-width: 5; text-align:center"></td>
      		<td class="fundo5" width="110" height="30" align="center"><input type="submit" value="Procurar"></td>
      	</form>
    </tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">		  
	<tr>
		<form action="consultas_completa2.asp"><td align="center" width="750"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;"></td></form>
	</tr>
</table>
<!--#include file="rodape.asp"--></center></body></html>
<% Else
	Response.Redirect("admin.asp")
End If %>