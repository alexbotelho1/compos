<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% If Session("LoggedIn") = True Then

	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from os order by os_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic
			contreg = 0
			Do While Not rsbanco.EOF
				If rsbanco("os_status") = 0 then
					contreg = contreg + 1					
				End IF
				rsbanco.movenext	
			Loop
		
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
			rsbanco.open "Select * from os order by os_solicdata ASC",banco,AdOpenKeySet,AdLockOptimistic
	End If
	If Request.QueryString("Ordem") = "4" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_solicdata DESC",banco,AdOpenKeySet,AdLockOptimistic
	End If	
	If Request.QueryString("Ordem") = "5" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_solicperiferico ASC",banco,AdOpenKeySet,AdLockOptimistic
	End If	
	If Request.QueryString("Ordem") = "6" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_solicperiferico DESC",banco,AdOpenKeySet,AdLockOptimistic
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
			rsbanco.open "Select * from os order by os_solicramal ASC",banco,AdOpenKeySet,AdLockOptimistic
	End If	
	If Request.QueryString("Ordem") = "14" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_solicramal DESC",banco,AdOpenKeySet,AdLockOptimistic
	End If
	If Request.QueryString("Ordem") = "15" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_status ASC",banco,AdOpenKeySet,AdLockOptimistic
	End If	
	If Request.QueryString("Ordem") = "16" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_status DESC",banco,AdOpenKeySet,AdLockOptimistic
	End If
		
	rsbanco.PageSize = registros
	contpage = (contreg \ registros) + 1 %>
<body><center>
<table border="1" width="700" height="102">
    <tr>
      <td class="fundo1" width="90" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td width="610" height="102" align="center" class="fundo2">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Controle de Ordem de Serviço</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Consulta das solicitações de abertura de Ordem de Serviço</b></font></td>
    </tr>
</table>
<% If (rsbanco.BOF And rsbanco.EOF) Or rsbanco.PageCount = 0 Then %>
<table border="0" cellpadding="0" cellspacing="0" width="100%" align="center">		  
	<tr>
		<td align="center"><font face="Trebuchet MS" size="2" color="#ffffff"><i>N&atilde;o h&aacute; nada no momento!</i></font></td>
	</tr>
</table>
<form action="cad_os.asp"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;"></form>
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
<table border="0" cellpadding="0" cellspacing="0" width="700" align="center">		  
	<tr>
		<td align="center" width="200">
		<% If rsbanco.AbsolutePage > 1 Then %>
			<a href="consultas3.asp?Ordem=<% = Disposicao %>&page=<% = (rsbanco.AbsolutePage - 1) %>"><img src="paganterior2.gif" align="absmiddle" border="0"></a>
		<% Else %>
			<img src="paganterior1.gif" align="absmiddle">
		<% End If %>
		</td>
		<td align="center" width="300"><font face="Trebuchet MS" size="2" color="#FFFFFF"><b>Esses são os <% = (rsbanco.PageSize * rsbanco.AbsolutePage) %> primeiros</b></font></td>
		<td align="center" width="200">
		<% If rsbanco.AbsolutePage < contpage Then %>
      		<a href="consultas3.asp?Ordem=<% = Disposicao %>&page=<% = (rsbanco.AbsolutePage + 1) %>"><img src="pagseguinte2.gif" align="absmiddle" border="0"></a>
   		<% 	Else %>
          	<img src="pagseguinte1.gif" align="absmiddle">
    	<% 	End If %>   			    		
       	</td>
	</tr>
</table>
<% Else %>
<table border="0" cellpadding="0" cellspacing="0" width="700" align="center">		  
	<tr>
		<td align="center" width="200">
		<% If rsbanco.AbsolutePage > 1 Then %>
			<a href="consultas3.asp?Ordem=<% = Disposicao %>&page=<% = (rsbanco.AbsolutePage - 1) %>"><img src="paganterior2.gif" align="absmiddle" border="0"></a>
		<% Else %>
			<img src="paganterior1.gif" align="absmiddle">
		<% End If %>
		</td>
		<td align="center" width="300"><font face="Trebuchet MS" size="2" color="#FFFFFF"><b>Registros do <% = ((rsbanco.PageSize * rsbanco.AbsolutePage) - 10) %>° até <% = (rsbanco.PageSize * rsbanco.AbsolutePage) %>°</b></font></td>
		<td align="center" width="200">
		<% If rsbanco.AbsolutePage < contpage Then %>
      		<a href="consultas3.asp?Ordem=<% = Disposicao %>&page=<% = (rsbanco.AbsolutePage + 1) %>"><img src="pagseguinte2.gif" align="absmiddle" border="0"></a>
   		<% 	Else %>
          	<img src="pagseguinte1.gif" align="absmiddle">
    	<% 	End If %>   			    		
       	</td>
	</tr>
</table>
<% End If
	pagina = rsbanco.AbsolutePage
	contador = contpage
%>
<!--#include file="statusos.asp"-->
<table width="700" align="center" border="1">
<tr>
	  <td class="fundo1" width="50" height="14" align="center"><a href="consultas3.asp?Ordem=1&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta.gif" alt="Crescente"></a>&nbsp;&nbsp;<a href="consultas3.asp?Ordem=2&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta2.gif" alt="Decrescente"></a></td>
      <td class="fundo1" width="120" height="14" align="center"><a href="consultas3.asp?Ordem=3&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta.gif" alt="Crescente"></a>&nbsp;&nbsp;<a href="consultas3.asp?Ordem=4&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta2.gif" alt="Decrescente"></a></td>
	  <td class="fundo1" width="90" height="14" align="center"><a href="consultas3.asp?Ordem=5&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta.gif" alt="Crescente"></a>&nbsp;&nbsp;<a href="consultas3.asp?Ordem=6&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta2.gif" alt="Decrescente"></a></td>
	  <td class="fundo1" width="140" height="14" align="center"><a href="consultas3.asp?Ordem=7&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta.gif" alt="Crescente"></a>&nbsp;&nbsp;<a href="consultas3.asp?Ordem=8&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta2.gif" alt="Decrescente"></a></td>
      <td class="fundo1" width="70" height="14" align="center"><a href="consultas3.asp?Ordem=9&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta.gif" alt="Crescente"></a>&nbsp;&nbsp;<a href="consultas3.asp?Ordem=10&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta2.gif" alt="Decrescente"></td>
      <td class="fundo1" width="140" height="14" align="center"><a href="consultas3.asp?Ordem=11&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta.gif" alt="Crescente"></a>&nbsp;&nbsp;<a href="consultas3.asp?Ordem=12&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta2.gif" alt="Decrescente"></td>
      <td class="fundo1" width="50" height="14" align="center"><a href="consultas3.asp?Ordem=13&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta.gif" alt="Crescente"></a>&nbsp;&nbsp;<a href="consultas3.asp?Ordem=14&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta2.gif" alt="Decrescente"></td>
	  <td class="fundo1" width="40" height="14" align="center"><a href="consultas3.asp?Ordem=15&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta.gif" alt="Crescente"></a>&nbsp;&nbsp;<a href="consultas3.asp?Ordem=16&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta2.gif" alt="Decrescente"></td>
	  <td class="fundo1" width="50" height="14" align="center">&nbsp;</td> 	    
</tr>
<tr>
	  <td class="fundo1" width="50" height="29" align="center"><font size="2"><b>N°</b></font></td>
      <td class="fundo1" width="120" height="29" align="center"><font size="2"><b>Data</b></font></td>
      <td class="fundo1" width="90" height="29" align="center"><b><font size="2">Periférico</td>
      <td class="fundo1" width="140" height="29" align="center"><b><font size="2">Solicitante</td>
      <td class="fundo1" width="70" height="29" align="center"><b><font size="2">Esquadrão</td>
      <td class="fundo1" width="140" height="29" align="center"><b><font size="2">Seção</td>
      <td class="fundo1" width="50" height="29" align="center"><b><font size="2">Ramal</td>
	  <td class="fundo1" width="40" height="29" align="center"><b><font size="2">Status</td>
	  <td class="fundo1" width="40" height="29" align="center"><b><font size="2">Abrir</td>      
</tr>
<% 	pagina1 = Request.QueryString("page")
 	If pagina1 > 1 then
		pulacont = 0
		Do While Not rsbanco.EOF And pulacont < registros
			If rsbanco("os_status") = 0 then
				pulacont = pulacont + 1
			End IF
			rsbanco.movenext
		Loop
	End If
	HowMany = 0
	Do While Not rsbanco.EOF And HowMany < rsbanco.PageSize
	If rsbanco("os_status") = 0 Then %>
<tr>
	  <td class="fundo5" width="50" height="29" align="center"><a href="solicitacao.asp?codsolic=<% = rsbanco("os_codigo") %>"><font color="#7F0D11" size="2"><b><% = rsbanco("os_codigo") %></a></td>
      <td class="fundo5" width="120" height="29" align="center"><font size="2"><% = rsbanco("os_solicdata") %></td>
      <td class="fundo5" width="90" height="29" align="center"><font size="2"><% = rsbanco("os_solicperiferico") %></td>
      <td class="fundo5" width="140" height="29" align="center"><font size="2"><% = rsbanco("os_solicmilitar") %></td>
      <td class="fundo5" width="70" height="29" align="center"><font size="2"><% = rsbanco("os_solicesquadrao") %></td>
      <td class="fundo5" width="140" height="29" align="center"><font size="2"><% = rsbanco("os_solicsecao") %></td>
      <td class="fundo5" width="50" height="29" align="center"><font size="2"><% = rsbanco("os_solicramal") %></td>
<% If rsbanco("os_status") = "1" then %>
	  <td class="fundo5" width="40" height="29" align="center"><img border="0" src="bolaverde.gif"></td>      
<% Else
		If rsbanco("os_status") = 2 then %>
      <td class="fundo5" width="40" height="29" align="center"><img border="0" src="bolaamarela.gif"></td>
		<% Else
			If rsbanco("os_status") = 3 then %>
      <td class="fundo5" width="40" height="29" align="center"><img border="0" src="bolaazul.gif"></td>		
			<% Else %>
      <td class="fundo5" width="40" height="29" align="center"><img border="0" src="bolavermelha.gif"></td>
			<% End If
		End If
	End If %>
	  <td class="fundo5" width="50" height="29" align="center"><font size="2"><a href="adicionar2.asp?codsolic=<% = rsbanco("os_codigo") %>"><img border="0" src="add.gif"></a></td>	
</tr>
<%		HowMany = HowMany + 1
		rsbanco.movenext
 	Else	
		rsbanco.movenext
	End If						 
	Loop %>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="700" align="center">		  
	<tr>
		<td align="center" width="200">
		<% If pagina > 1 Then %>
			<a href="consultas3.asp?Ordem=<% = Disposicao %>&page=<% = (pagina - 1) %>"><img src="paganterior2.gif" align="absmiddle" border="0"></a>
		<% Else %>
			<img src="paganterior1.gif" align="absmiddle">
		<% End If %>
		</td>
		<form action="cad_os.asp"><td align="center" width="300"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;"></td></form>
		<td align="center" width="200">
		<% If pagina < contador Then %>
      		<a href="consultas3.asp?Ordem=<% = Disposicao %>&page=<% = (pagina + 1) %>"><img src="pagseguinte2.gif" align="absmiddle" border="0"></a>
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