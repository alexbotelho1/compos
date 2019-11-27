<meta http-equiv="refresh" content="180"><html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!-- #include file="config.asp" -->
<!--#include file="styles.asp"-->
<% If Session("LoggedIn") = True Then
	If Request.QueryString("Ordem") = "" or Request.QueryString("Ordem") = "1" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic
	End If
	If Request.QueryString("Ordem") = "2" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_codigo DESC",banco,AdOpenKeySet,AdLockOptimistic
	End If
	If Request.QueryString("Ordem") = "3" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_ip ASC",banco,AdOpenKeySet,AdLockOptimistic
	End If
	If Request.QueryString("Ordem") = "4" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_ip DESC",banco,AdOpenKeySet,AdLockOptimistic
	End If	
	If Request.QueryString("Ordem") = "5" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_host ASC",banco,AdOpenKeySet,AdLockOptimistic
	End If	
	If Request.QueryString("Ordem") = "6" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_host DESC",banco,AdOpenKeySet,AdLockOptimistic
	End If	
	If Request.QueryString("Ordem") = "7" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_solicmilitar ASC",banco,AdOpenKeySet,AdLockOptimistic
	End If
	If Request.QueryString("Ordem") = "8" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_solicmilitar DESC",banco,AdOpenKeySet,AdLockOptimistic
	End If
	If Request.QueryString("Ordem") = "9" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_solicesquadrao ASC",banco,AdOpenKeySet,AdLockOptimistic
	End If	
	If Request.QueryString("Ordem") = "10" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_solicesquadrao DESC",banco,AdOpenKeySet,AdLockOptimistic
	End If	
	If Request.QueryString("Ordem") = "11" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_solicsecao ASC",banco,AdOpenKeySet,AdLockOptimistic
	End If	
	If Request.QueryString("Ordem") = "12" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_solicsecao DESC",banco,AdOpenKeySet,AdLockOptimistic
	End If
	If Request.QueryString("Ordem") = "13" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_logon ASC",banco,AdOpenKeySet,AdLockOptimistic
	End If	
	If Request.QueryString("Ordem") = "14" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_logon DESC",banco,AdOpenKeySet,AdLockOptimistic
	End If
	If Request.QueryString("Ordem") = "15" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_server ASC",banco,AdOpenKeySet,AdLockOptimistic
	End If	
	If Request.QueryString("Ordem") = "16" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_server DESC",banco,AdOpenKeySet,AdLockOptimistic
	End If
	If Request.QueryString("Ordem") = "17" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_numero ASC",banco,AdOpenKeySet,AdLockOptimistic
	End If	
	If Request.QueryString("Ordem") = "18" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_numero DESC",banco,AdOpenKeySet,AdLockOptimistic
	End If		
		
	rsbanco.PageSize = registros %>
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
	  </script></b></font></td>
    </tr>
</table>
<% If (rsbanco.BOF And rsbanco.EOF) Or rsbanco.PageCount = 0 Then %>
<table border="0" cellpadding="0" cellspacing="0" width="780" align="center">		  
	<tr>
		<td align="center"><font face="Trebuchet MS" size="2" color="#000000"><i>N&atilde;o h&aacute; nada no momento!</i></font></td>
	</tr>
</table>
<form action="cad_sol.asp"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;"></form>
<% Else
		If Request.QueryString("Ordem") = "" Then
			Disposicao = 1
		Else 
			Disposicao = Request.QueryString("Ordem")
		End If
		If Request.QueryString("page") > 0 Then  			   
			rsbanco.AbsolutePage = Request.QueryString("page")					
		Else				
			rsbanco.AbsolutePage = 1					
		End If 
if Request.QueryString("page") = 1 Then %>	
<table border="0" cellpadding="0" cellspacing="0" width="780" align="center">		  
	<tr>
		<td align="center" width="200">
		<% If rsbanco.AbsolutePage > 1 Then %>
			<a href="auditoria.asp?Ordem=<% = Disposicao %>&page=<% = (rsbanco.AbsolutePage - 1) %>"><img src="paganterior2.gif" align="absmiddle" border="0"></a>
		<% Else %>
			<img src="paganterior1.gif" align="absmiddle">
		<% End If %>
		</td>
		<td align="center" width="380"><font face="Trebuchet MS" size="2" color="#FFFFFF"><b>Esses são os <% = (rsbanco.PageSize * rsbanco.AbsolutePage) %> primeiros</b></font></td>
		<td align="center" width="200">
		<% If rsbanco.AbsolutePage < rsbanco.PageCount Then %>
      		<a href="auditoria.asp?Ordem=<% = Disposicao %>&page=<% = (rsbanco.AbsolutePage + 1) %>"><img src="pagseguinte2.gif" align="absmiddle" border="0"></a>
   		<% 	Else %>
          	<img src="pagseguinte1.gif" align="absmiddle">
    	<% 	End If %>   			    		
       	</td>
	</tr>
</table>
<% Else %>
<table border="0" cellpadding="0" cellspacing="0" width="780" align="center">		  
	<tr>
		<td align="center" width="200">
		<% If rsbanco.AbsolutePage > 1 Then %>
			<a href="auditoria.asp?Ordem=<% = Disposicao %>&page=<% = (rsbanco.AbsolutePage - 1) %>"><img src="paganterior2.gif" align="absmiddle" border="0"></a>
		<% Else %>
			<img src="paganterior1.gif" align="absmiddle">
		<% End If %>
		</td>
		<td align="center" width="380"><font face="Trebuchet MS" size="2" color="#FFFFFF"><b>Registros do <% = ((rsbanco.PageSize * rsbanco.AbsolutePage) - 10) %>° até <% = (rsbanco.PageSize * rsbanco.AbsolutePage) %>°</b></font></td>
		<td align="center" width="200">
		<% If rsbanco.AbsolutePage < rsbanco.PageCount Then %>
      		<a href="auditoria.asp?Ordem=<% = Disposicao %>&page=<% = (rsbanco.AbsolutePage + 1) %>"><img src="pagseguinte2.gif" align="absmiddle" border="0"></a>
   		<% 	Else %>
          	<img src="pagseguinte1.gif" align="absmiddle">
    	<% 	End If %>   			    		
       	</td>
	</tr>
</table>
<% End If
	pagina = rsbanco.AbsolutePage
	contador = rsbanco.PageCount
%>
<table width="780" align="center" border="1" height="23">
<tr>
	  <td class="fundo1" width="40" height="14" align="center"><a href="auditoria.asp?Ordem=1&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta.gif" alt="Crescente"></a>&nbsp;&nbsp;<a href="auditoria.asp?Ordem=2&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta2.gif" alt="Decrescente"></a></td>
	  <td class="fundo1" width="100" height="14" align="center"><a href="auditoria.asp?Ordem=7&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta.gif" alt="Crescente"></a>&nbsp;&nbsp;<a href="auditoria.asp?Ordem=8&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta2.gif" alt="Decrescente"></a></td>
      <td class="fundo1" width="100" height="14" align="center"><a href="auditoria.asp?Ordem=9&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta.gif" alt="Crescente"></a>&nbsp;&nbsp;<a href="auditoria.asp?Ordem=10&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta2.gif" alt="Decrescente"></td>
      <td class="fundo1" width="100" height="14" align="center"><a href="auditoria.asp?Ordem=11&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta.gif" alt="Crescente"></a>&nbsp;&nbsp;<a href="auditoria.asp?Ordem=12&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta2.gif" alt="Decrescente"></td>
      <td class="fundo1" width="100" height="14" align="center"><a href="auditoria.asp?Ordem=3&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta.gif" alt="Crescente"></a>&nbsp;&nbsp;<a href="auditoria.asp?Ordem=4&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta2.gif" alt="Decrescente"></a></td>
	  <td class="fundo1" width="100" height="14" align="center"><a href="auditoria.asp?Ordem=5&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta.gif" alt="Crescente"></a>&nbsp;&nbsp;<a href="auditoria.asp?Ordem=6&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta2.gif" alt="Decrescente"></a></td>      
      <td class="fundo1" width="100" height="14" align="center"><a href="auditoria.asp?Ordem=13&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta.gif" alt="Crescente"></a>&nbsp;&nbsp;<a href="auditoria.asp?Ordem=14&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta2.gif" alt="Decrescente"></td>
	  <td class="fundo1" width="100" height="14" align="center"><a href="auditoria.asp?Ordem=15&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta.gif" alt="Crescente"></a>&nbsp;&nbsp;<a href="auditoria.asp?Ordem=16&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta2.gif" alt="Decrescente"></td>
	  <td class="fundo1" width="40" height="14" align="center"><a href="auditoria.asp?Ordem=17&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta.gif" alt="Crescente"></a>&nbsp;&nbsp;<a href="auditoria.asp?Ordem=18&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta2.gif" alt="Decrescente"></td> 	    
</tr>
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
	Do While Not rsbanco.EOF And HowMany < rsbanco.PageSize	%>
<tr>
	  <td class="fundo5" width="40" height="27" align="center"><a href="solicitacao.asp?codsolic=<% = rsbanco("os_codigo") %>"><font color="#7F0D11" size="2"><b><% = rsbanco("os_codigo") %></a></td>
      <td class="fundo5" width="100" height="27" align="center"><font size="2"><% = rsbanco("os_solicmilitar") %></td>
      <td class="fundo5" width="100" height="27" align="center"><font size="2"><% = rsbanco("os_solicesquadrao") %></td>
      <td class="fundo5" width="100" height="27" align="center"><font size="2"><% = rsbanco("os_solicsecao") %></td>
      <td class="fundo5" width="100" height="27" align="center"><font size="2"><% = rsbanco("os_ip") %></td>
      <td class="fundo5" width="100" height="27" align="center"><font size="2"><% = rsbanco("os_host") %></td>
      <td class="fundo5" width="100" height="27" align="center"><font size="2"><% = rsbanco("os_logon") %></td>
	  <td class="fundo5" width="100" height="27" align="center"><font size="2"><% = rsbanco("os_server") %></td>      
      <td class="fundo5" width="40" height="27" align="center"><% If rsbanco("os_numero") > 0 Then %><a href="os.asp?codsolic=<% = rsbanco("os_codigo") %>"><% End If %><font color="#7F0D11" size="2"><b><% = rsbanco("os_numero") %></td>	
</tr>
<% 		HowMany = HowMany + 1
		rsbanco.MoveNext 				 
	Loop %>
</table>
<table border="1" width="780" height="30">
    <tr>
      	<form method="GET" action="auditoria_solicitacao.asp">
      		<td class="fundo1" width="145" height="30" align="center"><b><font size="2">N° da Solicitação</td>
      		<td class="fundo3" width="130" height="30" align="center">
            <input type="text" name="pesquisa_sol" size="10" style="border-style: inset; border-width: 5; background-color:#FFFF00; text-align:center"></td>
      		<td class="fundo5" width="115" height="30" align="center"><input type="submit" value="Procurar"></td>
      	</form>
      	<form method="GET" action="auditoria_ordem.asp">
      		<td class="fundo1" width="145" height="30" align="center"><b><font size="2">N° Ordem de Serviço</td>
      		<td class="fundo3" width="130" height="30" align="center"><input type="text" name="pesquisa_os" size="10" style="border-style: inset; border-width: 5; background-color:#00FFFF; text-align:center"></td>
      		<td class="fundo5" width="115" height="30" align="center"><input type="submit" value="Procurar"></td>
      	</form>
    </tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="780" align="center">		  
	<tr>
		<td align="center" width="200">
		<% If pagina > 1 Then %>
			<a href="auditoria.asp?Ordem=<% = Disposicao %>&page=<% = (pagina - 1) %>"><img src="paganterior2.gif" align="absmiddle" border="0"></a>
		<% Else %>
			<img src="paganterior1.gif" align="absmiddle">
		<% End If %>
		</td>
		<form action="admin.asp"><td align="center" width="380"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;"></td></form>
		<td align="center" width="200">
		<% If pagina < contador Then %>
      		<a href="auditoria.asp?Ordem=<% = Disposicao %>&page=<% = (pagina + 1) %>"><img src="pagseguinte2.gif" align="absmiddle" border="0"></a>
   		<% 	Else %>
          	<img src="pagseguinte1.gif" align="absmiddle">
    	<% 	End If %>    			    		
       	</td>
	</tr>
</table>
<!--#include file="rodape.asp"-->
<% End If %>
</center></body></html>
<% Else
	Response.Redirect("admin.asp")
End If %>