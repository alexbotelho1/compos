<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% If Session("LoggedIn") = True Then
	codsolic = request.querystring("codsolic")

	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from os WHERE os_codigo = "&codsolic&"",banco,AdOpenKeySet,AdLockOptimistic %>
<body><center><form method="GET" action="add4.asp">
<table border="1" width="700" height="102">
    <tr>
      <td class="fundo1" width="90" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="610" height="102" align="center">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Controle de Ordem de Serviço</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Consulta das solicitações de abertura de Ordem de Serviço</b></font></td>
    </tr>
</table>
<table width="700" align="center" border="1">
	<tr class="fundo1">
	  <td width="50" height="29" align="center"><b><font size="2">Número</td>
      <td width="120" height="29" align="center"><b><font size="2">Data</td>
      <td width="90" height="29" align="center"><b><font size="2">Periférico</td>
      <td width="140" height="29" align="center"><b><font size="2">Solicitante</td>
      <td width="70" height="29" align="center"><b><font size="2">Esquadrão</td>
      <td width="140" height="29" align="center"><b><font size="2">Seção</td>
      <td width="50" height="29" align="center"><b><font size="2">Ramal</td>
	  <td width="40" height="29" align="center"><b><font size="2">Status</td>      
	</tr>
	<tr class="fundo5">
	  <td width="50" height="29" align="center"><a href="solicitacao.asp?codsolic=<% = rsbanco("os_codigo") %>"><font color="#7F0D11" size="2"><b><% = rsbanco("os_codigo") %></a></td>      
      <td width="120" height="29" align="center"><font size="2"><% = rsbanco("os_solicdata") %></td>
      <td width="90" height="29" align="center"><font size="2"><% = rsbanco("os_solicperiferico") %></td>
      <td width="140" height="29" align="center"><font size="2"><% = rsbanco("os_solicmilitar") %></td>
      <td width="70" height="29" align="center"><font size="2"><% = rsbanco("os_solicesquadrao") %></td>
      <td width="140" height="29" align="center"><font size="2"><% = rsbanco("os_solicsecao") %></td>
      <td width="50" height="29" align="center"><font size="2"><% = rsbanco("os_solicramal") %></td>
<% If rsbanco("os_status") = 1 then %>
	  <td width="40" height="29" align="center"><img border="0" src="bolaverde.gif"></td>      
<% Else
		If rsbanco("os_status") = 2 then %>
      <td width="40" height="29" align="center"><img border="0" src="bolaamarela.gif"></td>
		<% Else
			If rsbanco("os_status") = 3 then %>
      <td width="40" height="29" align="center"><img border="0" src="bolaazul.gif"></td>		
			<% Else %>
      <td width="40" height="29" align="center"><img border="0" src="bolavermelha.gif"></td>
			<% End If
		End If
	End If %>	  	
	</tr>
</table>
<table width="700" align="center" border="1">
	<tr class="fundo1">
	  <td width="80" height="29" align="center"><b><font size="2">Número OS</td>
      <td width="120" height="29" align="center"><b><font size="2">Data</td>
      <td width="330" height="29" align="center"><b><font size="2">Observação</td>
      <td width="140" height="29" align="center"><b><font size="2">Militar STI</td>
      <td width="50" height="29" align="center"><b><font size="2">Ramal</td>      
	</tr>
	<tr class="fundo5">
	  <td width="80" height="29" align="center"><a href="os.asp?codsolic=<% = rsbanco("os_codigo") %>"><font color="#7F0D11" size="2"><b><% = rsbanco("os_numero") %></a></td>
      <td width="120" height="29" align="center"><font size="2"><% = rsbanco("os_dataaber") %></td>
      <td width="330" height="29" ><font size="2"><% = rsbanco("os_descricaoaber") %></td>
      <td width="140" height="29" align="center"><font size="2"><% = rsbanco("os_militaraber") %></td>
      <td width="50" height="29" align="center"><font size="2"><% = rsbanco("os_ramalaber") %></td>	  
	</tr>
</table>
<table width="700" align="center" border="1">
	<tr class="fundo1">
	  <td width="50" height="29" align="center"><b><font size="2">Tempo</td>
      <td width="100" height="29" align="center"><b><font size="2">Data Exec</td>
      <td width="300" height="29" align="center"><b><font size="2">Sv Executado</td>
      <td width="190" height="29" align="center"><b><font size="2">Mat Utilizado</td>
      <td width="80" height="29" align="center"><b><font size="2">Militar Exec</td>      
	</tr>
	<tr class="fundo5">
	  <td width="50" height="29" align="center"><font size="2"><b><% = rsbanco("os_numero") %></td>
      <td width="100" height="29" align="center"><font size="2"><% = rsbanco("os_dataexec") %></td>
      <td width="300" height="29" ><font size="1"><% = rsbanco("os_descricaoexec") %></td>
      <td width="190" height="29" align="center"><font size="1"><% = rsbanco("os_matusadoexec") %></td>
      <td width="80" height="29" align="center"><font size="2"><% = rsbanco("os_militarexec") %></font></td>	  
	</tr class="fundo1">
</table>
<input type=hidden name=os_codigo value="<%=rsbanco("os_codigo")%>">
<input type=hidden name=os_status value="<%=rsbanco("os_status")%>">
<table border="1" width="700" height="50">
	<tr>
        <td class="fundo1" width="140" height="50" align="center"><b>Data Fechamento</b></td>
        <td class="fundo3" width="90" height="50"><input readonly name="os_dataconc" type="text" value="<% Response.write Now %>" size="20" style="text-align: center"></td>
        <td class="fundo1" width="140" height="50" align="center"><b>Militar Fechou</b></td>
        <td class="fundo3" width="95" height="50">
	<% 	Set rsbanco2 = Server.CreateObject("ADODB.Recordset")
		rsbanco2.Open "SELECT * FROM sti ORDER BY sti_antiguidade ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="os_militarconc">					    
		<%  Do While Not rsbanco2.EOF %>								    
				<option value="<% = rsbanco2("sti_nomeguerra") %>"><% = rsbanco2("sti_nomeguerra") %></option>
		<% 		rsbanco2.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco2.Close
		Set rsbanco2= Nothing %>   
  		</td>
        <td class="fundo1" width="140" height="50" align="center"><b>Militar Recebeu</b></td>
        <td class="fundo3" width="95" height="50"><input type="text" name="os_milrecconc" size="20"></td>
	</tr>
	<tr>
        <td class="fundo1" width="140" height="120" align="center"><b>Observação</b><p><b>de</b></p><p><b>Fechamento</b></td>
        <td class="fundo3" width="560" height="120" colspan="5"><textarea name="os_observconc" cols="72" rows="8">Nada a relatar</textarea></td>
	</tr>            
</table>
<input type="submit" value="Salvar" name="BTincluir"><input type="reset" value="Limpar" name="BTlimpar"></form>
<form action="consultasos2.asp"> <input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;">
<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font color="#FFFFFF" size="2">OBS: Todos os Campos são obrigatórios. Verifique as informações antes de salvá-las,</font></b></p>
<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font color="#FFFFFF" size="2">&nbsp; pois as mesmas não poderão ser apagadas pelo usuário solicitante.</font></b></p>
<!--#include file="rodape.asp"--></form></center></body></html>
<% Else
	Response.Redirect("admin.asp")
End If %>