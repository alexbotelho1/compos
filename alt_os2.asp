<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% If Session("LoggedIn") = True Then	
	codsolic = request.querystring("codsolic")

	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from os WHERE os_codigo = "&codsolic&"",banco,AdOpenKeySet,AdLockOptimistic %>
<body><center><form method="GET" action="alt9.asp">
<table border="1" width="600" height="102">
    <tr>
      <td class="fundo1" width="100" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="500" height="102" align="center">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Controle de Ordem de Serviço</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Editar Ordem de Serviço</b></font></td>
    </tr>
</table>
<table border="1" width="600" height="23">
      <tr>
        <td class="fundo2" width="600" height="23" align="center" colspan="4"><font color="#008000"><b>Formulário de Solicitação</b></font></td>
      </tr>
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>Número</b></td>
        <td class="fundo3" width="500" height="23" colspan="3"><%=rsbanco("os_codigo")%></td>
      </tr>
		  <input type=hidden name=os_codigo value="<%=rsbanco("os_codigo")%>">        
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>Data Solic</b></td>
        <td class="fundo3" width="150" height="23"><input readonly value="<%=rsbanco("os_solicdata")%>" name="os_solicdata" type="text" size="20" style="text-align: center"></td>
        <td class="fundo1" width="100" height="23" align="center"><b>Periférico</b></td>
        <td class="fundo3" width="150" height="23">
	<% 	Set rsbanco5 = Server.CreateObject("ADODB.Recordset")
		rsbanco5.Open "SELECT * FROM periferico ORDER BY periferico_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="os_solicperiferico">					    
		<%  Do While Not rsbanco5.EOF %>								    
				<option value="<% = rsbanco5("periferico_nome") %>"<% If rsbanco("os_solicperiferico") = rsbanco5("periferico_nome") Then Response.Write (" selected") %>><% = rsbanco5("periferico_nome") %></option>
		<% 		rsbanco5.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco5.Close
		Set rsbanco5= Nothing %>         
        </td>
      </tr>      
      <tr>
        <td class="fundo1" width="100" height="197" align="center"><b>Descrição</b></p><p align="center"><b>do</b></p><p align="center"><b>Problema</b></td>
        <td class="fundo3" width="500" height="197" colspan="3"><textarea name="os_solicdescricao" cols="59" rows="12"><%=rsbanco("os_solicdescricao")%></textarea></td>
      </tr>
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>Solicitante</b></td>
        <td class="fundo3" width="150" height="23"><input value="<%=rsbanco("os_solicmilitar")%>" type="text" name="os_solicmilitar" size="20"></td>
        <td class="fundo1" width="100" height="23" align="center"><b>Esquadrão</b></td>
        <td class="fundo3" width="150" height="23">
	<% 	Set rsbanco4 = Server.CreateObject("ADODB.Recordset")
		rsbanco4.Open "SELECT * FROM esquadrao ORDER BY esquadrao_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="os_solicesquadrao">					    
		<%  Do While Not rsbanco4.EOF %>								    
				<option value="<% = rsbanco4("esquadrao_nome") %>"<% If rsbanco("os_solicesquadrao") = rsbanco4("esquadrao_nome") Then Response.Write (" selected") %>><% = rsbanco4("esquadrao_nome") %></option>
		<% 		rsbanco4.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco4.Close
		Set rsbanco4= Nothing %>        
		</td>
      </tr>
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>Seção</b></td>
        <td class="fundo3" width="150" height="23">
	<% 	Set rsbanco3 = Server.CreateObject("ADODB.Recordset")
		rsbanco3.Open "SELECT * FROM secao ORDER BY secao_nome ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="os_solicsecao">					    
		<%  Do While Not rsbanco3.EOF %>								    
				<option value="<% = rsbanco3("secao_nome") %>"<% If rsbanco("os_solicsecao") = rsbanco3("secao_nome") Then Response.Write (" selected") %>><% = rsbanco3("secao_nome") %></option>
		<% 		rsbanco3.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco3.Close
		Set rsbanco3= Nothing %>
        </td>
        <td class="fundo1" width="100" height="23" align="center"><b>Ramal</b></td>
        <td class="fundo3" width="150" height="23"><input value="<%=rsbanco("os_solicramal")%>" type="text" name="os_solicramal" size="12"></td>
      </tr>       
    </table>
    <table border="1" width="600" height="23">
      <tr>
        <td class="fundo2" width="600" height="23" align="center" colspan="4"><font color="#008000"><b>Formulário de Abertura de Ordem de Serviço</b></font></td>
      </tr>    
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>OS Número</b></td>     
        <td class="fundo3" width="150" height="23"><input value="<%=rsbanco("os_numero")%>" name="os_numero" type="text" size="8" style="text-align: center"></td>
        <td class="fundo1" width="100" height="23" align="center"><b>OS Data</b></td>
        <td class="fundo3" width="150" height="23"><input readonly value="<%=rsbanco("os_dataaber")%>" name="os_dataaber" type="text" size="20" style="text-align: center"></td>
      </tr>    
      <tr>
        <td class="fundo1" width="100" height="120" align="center"><b>Observações</b></td>
        <td class="fundo3" width="600" height="120" colspan="3">
        <textarea name="os_descricaoaber" cols="59" rows="12"><%=rsbanco("os_descricaoaber")%></textarea></td>
      </tr>
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>Militar STI</b></td>
        <td class="fundo3" width="150" height="23">
	<% 	Set rsbanco2 = Server.CreateObject("ADODB.Recordset")
		rsbanco2.Open "SELECT * FROM sti ORDER BY sti_antiguidade ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="os_militaraber">					    
		<%  Do While Not rsbanco2.EOF %>								    
				<option value="<% = rsbanco2("sti_nomeguerra") %>" <% If rsbanco("os_militaraber") = rsbanco2("sti_nomeguerra") Then Response.Write (" selected") %>><% = rsbanco2("sti_nomeguerra") %></option>
		<% 		rsbanco2.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco2.Close
		Set rsbanco2 = Nothing %>        
        <td class="fundo1" width="100" height="23" align="center"><b>Ramal STI</b></td>
        <td class="fundo3" width="150" height="23"><input value="<%=rsbanco("os_ramalaber")%>" type="text" value="9714" name="os_ramalaber" size="8" style="text-align: center"></td>
      </tr>       
    </table> 
    <table border="1" width="600" height="23">
      <tr>
        <td class="fundo2" width="600" height="23" align="center" colspan="4"><font color="#008000"><b>Formulário de Execução de Ordem de Serviço</b></font></td>
      </tr>     
      <tr>
        <td class="fundo1" width="130" height="23" align="center"><b>Tempo H/H</b></td>
        <td class="fundo3" width="155" height="23"><input value="<%=rsbanco("os_tempoexec")%>" name="os_tempoexec" type="text" size="20" style="text-align: center"></td>
        <td class="fundo1" width="160" height="23" align="center"><b>Data Exec.</b></td>
        <td class="fundo3" width="155" height="23"><input readonly value="<%=rsbanco("os_dataexec")%>" name="os_dataexec" type="text" size="20" style="text-align: center"></td>
      </tr>    
      <tr>
        <td class="fundo1" width="130" height="134" align="center"><b>Serviço</b><p align="center"><b>Executado</b></td>
        <td class="fundo3" width="470" height="134" colspan="3"><textarea name="os_descricaoexec" cols="60" rows="8"><%=rsbanco("os_descricaoexec")%></textarea></td>
      </tr>
      <tr>
        <td class="fundo1" width="130" height="121" align="center"><b>Material</b><p align="center"><b>Utilizado</b></td>
        <td class="fundo3" width="470" height="121" colspan="3"><textarea name="os_matusadoexec" cols="60" rows="7"><%=rsbanco("os_matusadoexec")%></textarea></td>
      </tr>      
      <tr>
        <td class="fundo1" width="130" height="23" align="center"><b>Militar Exec.</b></td>
        <td class="fundo3" width="470" height="23" colspan="3">
	<% 	Set rsbanco2 = Server.CreateObject("ADODB.Recordset")
		rsbanco2.Open "SELECT * FROM sti ORDER BY sti_antiguidade ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="os_militarexec">					    
		<%  Do While Not rsbanco2.EOF %>								    
				<option value="<% = rsbanco2("sti_nomeguerra") %>" <% If rsbanco("os_militarexec") = rsbanco2("sti_nomeguerra") Then Response.Write (" selected") %>><% = rsbanco2("sti_nomeguerra") %></option>
		<% 		rsbanco2.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco2.Close
		Set rsbanco2 = Nothing %>                  
        </td>
      </tr>       
    </table>      
<input type="submit" value="Salvar" name="BTincluir">&nbsp;&nbsp;&nbsp;&nbsp;<input type="reset" value="Limpar" name="BTlimpar"></form>
<form action="consultasos2.asp"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;">
<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font color="#FFFFFF" size="2">OBS: Todos os Campos são obrigatórios. Verifique as informações antes de salvá-las,</font></b></p>
<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font color="#FFFFFF" size="2">&nbsp; pois as mesmas não poderão ser apagadas pelo usuário solicitante.</font></b></p>
<!--#include file="rodape.asp"--></form></center></body></html>
<% Else
	Response.Redirect("admin.asp")
End If %>