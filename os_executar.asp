<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% If Session("LoggedIn") = True Then
	codsolic = request.querystring("codsolic")

	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from os WHERE os_codigo = "&codsolic&"",banco,AdOpenKeySet,AdLockOptimistic %>
<body><center><form method="GET" action="add3.asp">
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
<input type=hidden name=os_codigo value="<%=rsbanco("os_codigo")%>">
<input type=hidden name=os_status value="<%=rsbanco("os_status")%>">
<table border="1" width="700" height="1">
	<tr>
        <td class="fundo1" width="270" height="23" align="center"><b>Tempo H/H</b></td>
        <td class="fundo3" width="80" height="23"><input name="os_tempoexec" type="text" size="20" style="text-align: center"></td>
        <td class="fundo1" width="200" height="23" align="center"><b>Data Exec.</b></td>
        <td class="fundo3" width="150" height="23"><input readonly name="os_dataexec" type="text" value="<% Response.write Now %>" size="20" style="text-align: center"></td>
      </tr>    
      <tr>
        <td class="fundo1" width="270" height="120" align="center"><b>Serviço</b><p align="center"><b>Executado</b></td>
        <td class="fundo3" width="430" height="120" colspan="3"><textarea name="os_descricaoexec" cols="70" rows="8">Nada a relatar</textarea></td>
      </tr>
      <tr>
        <td class="fundo1" width="270" height="120" align="center"><b>Material</b><p align="center"><b>Utilizado</b></td>
        <td class="fundo3" width="430" height="120" colspan="3"><textarea name="os_matusadoexec" cols="70" rows="7">Nada a relatar</textarea></td>
      </tr>      
      <tr>
        <td class="fundo1" width="270" height="23" align="center"><b>Militar Exec.</b></td>
        <td class="fundo3" width="430" height="23" colspan="3">
	<% 	Set rsbanco2 = Server.CreateObject("ADODB.Recordset")
		rsbanco2.Open "SELECT * FROM sti ORDER BY sti_antiguidade ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="os_militarexec">					    
		<%  Do While Not rsbanco2.EOF %>								    
				<option value="<% = rsbanco2("sti_nomeguerra") %>"><% = rsbanco2("sti_nomeguerra") %></option>
		<% 		rsbanco2.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco2.Close
		Set rsbanco2= Nothing %>         
  		</td>
	</tr>       
</table>
<input type="submit" value="Salvar" name="BTincluir"><input type="reset" value="Limpar" name="BTlimpar">
<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font color="#FFFFFF" size="2">OBS: Todos os Campos são obrigatórios. Verifique as informações antes de salvá-las,</font></b></p>
<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font color="#FFFFFF" size="2">&nbsp; pois as mesmas não poderão ser apagadas pelo usuário solicitante.</font></b></p>
</form><form action="consultasos2.asp"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;">
<!--#include file="rodape.asp"--></form></center></body></html>
<% Else
	Response.Redirect("admin.asp")
End If %>