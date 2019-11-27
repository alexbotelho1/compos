<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% If Session("LoggedIn") = True Then
	codigo = request.querystring("codigo")

	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from switch WHERE switch_codigo = "&codigo&"",banco,AdOpenKeySet,AdLockOptimistic
%>
<body><center><form method="GET" action="alt7.asp">
<table border="1" width="600" height="102">
    <tr>
      <td class="fundo1" width="86" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="514" height="102" align="center">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Inventário de Informática</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Alterar Switch</b></font></td>
    </tr>
</table>
<table border="1" width="599" height="1">
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Data</b></td>
        <td class="fundo3" width="210" height="23"><input name="switch_data" type="text" value="<% Response.write Now %>" size="20" style="border-style: inset; border-width: 5; text-align:center"></td>
        <td class="fundo1" width="90" height="23" align="center"><b>FCG</b></td>
        <td class="fundo3" width="210" height="23"><input name="switch_fcg" type="text" value="<% = rsbanco("switch_fcg") %>" size="12" style="border-style: inset; border-width: 5"></td>
      </tr>
			<input type=hidden name=switch_codigo value="<%=rsbanco("switch_codigo")%>">      
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Seção</b></td>
        <td class="fundo3" width="210" height="23">
	<% 	Set rsbanco3 = Server.CreateObject("ADODB.Recordset")
		rsbanco3.Open "SELECT * FROM secao ORDER BY secao_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="switch_secao">				    
		<%  Do While Not rsbanco3.EOF %>								    
				<option value="<% = rsbanco3("secao_nome") %>"<% If rsbanco("switch_secao") = rsbanco3("secao_nome") Then Response.Write (" selected") %>><% = rsbanco3("secao_nome") %></option>
		<% 		rsbanco3.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco3.Close
		Set rsbanco3= Nothing %>
        </td>
		<td class="fundo1" width="90" height="23" align="center"><b>Esquadrão</b></td>
        <td class="fundo3" width="210" height="23">
	<% 	Set rsbanco4 = Server.CreateObject("ADODB.Recordset")
		rsbanco4.Open "SELECT * FROM esquadrao ORDER BY esquadrao_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="switch_esquadrao">				    
		<%  Do While Not rsbanco4.EOF %>								    
				<option value="<% = rsbanco4("esquadrao_nome") %>"<% If rsbanco("switch_esquadrao") = rsbanco4("esquadrao_nome") Then Response.Write (" selected") %>><% = rsbanco4("esquadrao_nome") %></option>
		<% 		rsbanco4.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco4.Close
		Set rsbanco4= Nothing %>
        </td>
      </tr>      
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Marca</b></td>
        <td class="fundo3" width="210" height="23">
	<% 	Set rsbanco2 = Server.CreateObject("ADODB.Recordset")
		rsbanco2.Open "SELECT * FROM marcanb ORDER BY marcanb_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="switch_marca">	 
		<%  Do While Not rsbanco2.EOF %>								    
				<option value="<% = rsbanco2("marcanb_nome") %>"<% If rsbanco("switch_marca") = rsbanco2("marcanb_nome") Then Response.Write (" selected") %>><% = rsbanco2("marcanb_nome") %></option>
		<% 		rsbanco2.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco2.Close
		Set rsbanco2 = Nothing %>   
        </td>      
		<td class="fundo1" width="90" height="23" align="center"><b>Situação</b></td>
        <td class="fundo3" width="210" height="23">
            <select size="1" name="switch_situacao">
  				<option value="Uso"<% If rsbanco("switch_situacao") = "Uso" Then Response.Write (" selected") %>>Uso</option>
                <option value="Manutenção"<% If rsbanco("switch_situacao") = "Manutenção" Then Response.Write (" selected") %>>Manutenção</option>
                <option value="Sucata"<% If rsbanco("switch_situacao") = "Sucata" Then Response.Write (" selected") %>>Sucata</option> 
  			</select>         
        </td>
      </tr>
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Portas</b></td>
        <td class="fundo3" width="510" height="23" colspan="3"><input type="text" name="switch_porta" value="<% = rsbanco("switch_porta") %>"size="12" style="border-style: inset; border-width: 5"><b><font size="2" color="#070E5A">Quantidade de portas no Hub ou na Switch.</font></b></td>
      </tr>      
      <tr>
        <td class="fundo1" width="90" height="65" align="center"><b>Observação</b></td>
        <td class="fundo3" width="510" height="65" colspan="3"><textarea name="switch_observa" cols="57" rows="6" style="border-style: inset; border-width: 5"><% = rsbanco("switch_observa")%></textarea></td>
      </tr>                
</table>
<input type="submit" value="&nbsp;&nbsp;Salvar&nbsp;&nbsp;" name="BTincluir">
</form>
<form action="consultas4.asp"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;">
<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2" color="#FFFFFF">OBS: Todos os Campos são obrigatórios. Verifique as informações antes de salvá-las,</font></b></p>
<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2" color="#FFFFFF">&nbsp; pois as mesmas não poderão ser apagadas pelo usuário solicitante.</font></b></p>
<!--#include file="rodape.asp"--></form>
</center></body></html>
<% Else
	Response.Redirect("admin.asp")
End If %>