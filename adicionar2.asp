<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% If Session("LoggedIn") = True Then
	codsolic = request.querystring("codsolic")

	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from os WHERE os_codigo = "&codsolic&"",banco,AdOpenKeySet,AdLockOptimistic %>
<body><center><form method="GET" action="add2.asp">
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
	  <td width="50" height="29" align="center" ><b><font size="2">Número</td>
      <td width="120" height="29" align="center" ><b><font size="2">Data</td>
      <td width="90" height="29" align="center" ><b><font size="2">Periférico</td>
      <td width="140" height="29" align="center" ><b><font size="2">Solicitante</td>
      <td width="70" height="29" align="center" ><b><font size="2">Esquadrão</td>
      <td width="140" height="29" align="center" ><b><font size="2">Seção</td>
      <td width="50" height="29" align="center" ><b><font size="2">Ramal</td>
	  <td width="40" height="29" align="center" ><b><font size="2">Status</td>      
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
<input type=hidden name=os_codigo value="<%=rsbanco("os_codigo")%>">
<input type=hidden name=os_status value="<%=rsbanco("os_status")%>">
<table border="1" width="700" height="21">
	<tr>
        <td class="fundo1" width="100"  height="23" align="center"><b>OS Número</b></td>
<%	set rsbanco1=server.createobject("ADODB.Recordset")
		rsbanco1.open "Select * from os order by os_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic
	
	rsbanco1.MoveLast
		ultimo = rsbanco1("os_codigo")
	rsbanco1.MoveFirst
		numeroos = 1
	
	Do While rsbanco1("os_codigo") <> ultimo
		If rsbanco1("os_numero") > 0 then
			numeroos = numeroos + 1
			rsbanco1.MoveNext
		Else
			rsbanco1.MoveNext
		End If
	Loop %>        
        <td class="fundo3" width="250" height="23"><input readonly name="os_numero" type="text" value="<% = numeroos %>" size="8" style="text-align: center"></td>
        <td class="fundo1" width="100"  height="23" align="center"><b>OS Data</b></td>
        <td class="fundo3" width="250" height="23"><input readonly name="os_dataaber" type="text" value="<% Response.write Now %>" size="20" style="text-align: center"></td>
      </tr>    
      <tr>
        <td class="fundo1" width="100"  height="197" align="center"><p><b>Observações</b></td>
        <td class="fundo3" width="600" height="197" colspan="3"><textarea name="os_descricaoaber" cols="72" rows="12">Nada a relatar</textarea></td>
      </tr>
      <tr>
        <td class="fundo1" width="100" height="21" align="center"><b>Militar STI</b></td>
        <td class="fundo3" width="250" height="21">
	<% 	Set rsbanco2 = Server.CreateObject("ADODB.Recordset")
		rsbanco2.Open "SELECT * FROM sti ORDER BY sti_antiguidade ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="os_militaraber">					    
		<%  Do While Not rsbanco2.EOF %>								    
				<option value="<% = rsbanco2("sti_nomeguerra") %>"<% If rsbanco("os_militaraber") = rsbanco2("sti_nomeguerra") Then Response.Write (" selected") %>><% = rsbanco2("sti_nomeguerra") %></option>
		<% 		rsbanco2.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco2.Close
		Set rsbanco2= Nothing %>        
		</td>
        <td class="fundo1" width="100" height="21" align="center"><b>Ramal STI</b></td>
        <td class="fundo3" width="250" height="21"><input type="text" value="9714" name="os_ramalaber" size="8" style="text-align: center"></td>
	</tr>       
</table>
<input type="submit" value="Salvar" name="BTincluir"><input type="reset" value="Limpar" name="BTlimpar"></form>
<form action="consultas3.asp"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;">
<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font color="#FFFFFF" size="2">OBS: Todos os Campos são obrigatórios. Verifique as informações antes de salvá-las,</font></b>
<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font color="#FFFFFF" size="2">&nbsp; pois as mesmas não poderão ser apagadas pelo usuário solicitante.</font></b></p>
<!--#include file="rodape.asp"--></form></center></body></html>
<% Else
	Response.Redirect("admin.asp")
End If %>