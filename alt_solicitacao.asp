<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% If Session("LoggedIn") = True Then
	codsolic = request.querystring("codsolic")

	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from os WHERE os_codigo = "&codsolic&"",banco,AdOpenKeySet,AdLockOptimistic
%>
<body><center><form method="GET" action="alt.asp">
<table border="1" width="600" height="102">
    <tr>
      <td class="fundo1" width="90" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td width="510" height="102" align="center" class="fundo2">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Controle de Ordem de Serviço</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Consulta das solicitações de abertura de Ordem de Serviço</b></font></td>
    </tr>
</table>
<table border="1" width="600" height="23">
      <tr>
        <td class="fundo1" width="100" height="23"><b>Número</b></td>
        <td class="fundo3" width="500" height="23" colspan="3"><%=rsbanco("os_codigo")%></td>
      </tr> 
      <input type=hidden name=os_codigo value="<%=rsbanco("os_codigo")%>">
      <tr>
        <td class="fundo1" width="81" height="23" align="center"><b>Data</b></td>
        <td class="fundo3" width="91" height="23"><input readonly name="os_solicdata" type="text" value="<%=rsbanco("os_solicdata")%>" size="20" style="border-style: inset; border-width: 5; text-align:center"></td>
        <td class="fundo1" width="154" height="23" align="center"><b>Periférico</b></td>
        <td class="fundo3" width="521" height="23">
	<% 	Set rsbanco5 = Server.CreateObject("ADODB.Recordset")
		rsbanco5.Open "SELECT * FROM periferico ORDER BY periferico_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="os_solicperiferico">
				<option value="0" selected>Selecione</option>				    
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
        <td class="fundo1" width="81" height="146" align="center"><b>Descrição</b></p><p align="center"><b>do</b></p><p align="center"><b>Problema</b></td>
        <td class="fundo3" width="505" height="146" colspan="3"><textarea name="os_solicdescricao" cols="60" rows="9" style="border-style: inset; border-width: 5"><%=rsbanco("os_solicdescricao")%></textarea></td>
      </tr>
      <tr>
        <td class="fundo1" width="81" height="23" align="center"><b>Solicitante</b></td>
        <td class="fundo3" width="91" height="23"><input type="text" name="os_solicmilitar" value="<%=rsbanco("os_solicmilitar")%>" size="20" style="border-style: inset; border-width: 5"></td>
        <td class="fundo1" width="154" height="23" align="center"><b>Esquadrão</b></td>
        <td class="fundo3" width="202" height="23">
	<% 	Set rsbanco4 = Server.CreateObject("ADODB.Recordset")
		rsbanco4.Open "SELECT * FROM esquadrao ORDER BY esquadrao_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="os_solicesquadrao">
				<option value="0" selected>Selecione</option>						    
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
        <td class="fundo1" width="81" height="23" align="center"><b>Seção</b></td>
        <td class="fundo3" width="412" height="23">
	<% 	Set rsbanco3 = Server.CreateObject("ADODB.Recordset")
		rsbanco3.Open "SELECT * FROM secao ORDER BY secao_nome ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="os_solicsecao">
				<option value="0" selected>Selecione</option>			    
		<%  Do While Not rsbanco3.EOF %>								    
				<option value="<% = rsbanco3("secao_nome") %>"<% If rsbanco("os_solicsecao") = rsbanco3("secao_nome") Then Response.Write (" selected") %>><% = rsbanco3("secao_nome") %></option>
		<% 		rsbanco3.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco3.Close
		Set rsbanco3= Nothing %>  
        </td>
        <td class="fundo1" width="154" height="23" align="center"><b>Ramal</b></td>
        <td class="fundo3" width="523" height="23"><input type="text" name="os_solicramal" value="<%=rsbanco("os_solicramal")%>" size="12" style="border-style: inset; border-width: 5"></td>
      </tr>       
    </table>
<input type="submit" value="&nbsp;&nbsp;Salvar&nbsp;&nbsp;" name="BTincluir">
</form>
<form action="javascript:history.go(-1)"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;">
<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2" color="#FFFFFF">OBS: Todos os Campos são obrigatórios. Verifique as informações antes de salvá-las,</font></b></p>
<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2" color="#FFFFFF">&nbsp; pois as mesmas não poderão ser apagadas pelo usuário solicitante.</font></b></p>
<!--#include file="rodape.asp"--></form>
</center></body></html>
<% Else
	Response.Redirect("admin.asp")
End If %>