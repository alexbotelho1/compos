<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% If Session("LoggedIn") = True Then
	codigo = request.querystring("codigo")

	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from estabilizador WHERE estabilizador_codigo = "&codigo&"",banco,AdOpenKeySet,AdLockOptimistic
%>
<body><center><form method="GET" action="alt6.asp">
<table border="1" width="600" height="102">
    <tr>
      <td class="fundo1" width="86" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="514" height="102" align="center">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Invent�rio de Inform�tica</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Alterar Estabilizador</b></font></td>
    </tr>
</table>
<table border="1" width="599" height="1">
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Data</b></td>
        <td class="fundo3" width="210" height="23"><input name="estabilizador_data" type="text" value="<% Response.write Now %>" size="20" style="border-style: inset; border-width: 5; text-align:center"></td>
        <td class="fundo1" width="90" height="23" align="center"><b>FCG</b></td>
        <td class="fundo3" width="210" height="23"><input name="estabilizador_fcg" type="text" value="<% = rsbanco("estabilizador_fcg") %>" size="12" style="border-style: inset; border-width: 5"></td>
      </tr>
			<input type=hidden name=estabilizador_codigo value="<%=rsbanco("estabilizador_codigo")%>">      
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Se��o</b></td>
        <td class="fundo3" width="210" height="23">
	<% 	Set rsbanco3 = Server.CreateObject("ADODB.Recordset")
		rsbanco3.Open "SELECT * FROM secao ORDER BY secao_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="estabilizador_secao">				    
		<%  Do While Not rsbanco3.EOF %>								    
				<option value="<% = rsbanco3("secao_nome") %>"<% If rsbanco("estabilizador_secao") = rsbanco3("secao_nome") Then Response.Write (" selected") %>><% = rsbanco3("secao_nome") %></option>
		<% 		rsbanco3.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco3.Close
		Set rsbanco3= Nothing %>
        </td>
		<td class="fundo1" width="90" height="23" align="center"><b>Esquadr�o</b></td>
        <td class="fundo3" width="210" height="23">
	<% 	Set rsbanco4 = Server.CreateObject("ADODB.Recordset")
		rsbanco4.Open "SELECT * FROM esquadrao ORDER BY esquadrao_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="estabilizador_esquadrao">				    
		<%  Do While Not rsbanco4.EOF %>								    
				<option value="<% = rsbanco4("esquadrao_nome") %>"<% If rsbanco("estabilizador_esquadrao") = rsbanco4("esquadrao_nome") Then Response.Write (" selected") %>><% = rsbanco4("esquadrao_nome") %></option>
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
			<select name="estabilizador_marca">	 
		<%  Do While Not rsbanco2.EOF %>								    
				<option value="<% = rsbanco2("marcanb_nome") %>"<% If rsbanco("estabilizador_marca") = rsbanco2("marcanb_nome") Then Response.Write (" selected") %>><% = rsbanco2("marcanb_nome") %></option>
		<% 		rsbanco2.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco2.Close
		Set rsbanco2 = Nothing %>   
        </td>      
		<td class="fundo1" width="90" height="23" align="center"><b>Situa��o</b></td>
        <td class="fundo3" width="210" height="23">
            <select size="1" name="estabilizador_situacao">
  				<option value="Uso"<% If rsbanco("estabilizador_situacao") = "Uso" Then Response.Write (" selected") %>>Uso</option>
                <option value="Manuten��o"<% If rsbanco("estabilizador_situacao") = "Manuten��o" Then Response.Write (" selected") %>>Manuten��o</option>
                <option value="Sucata"<% If rsbanco("estabilizador_situacao") = "Sucata" Then Response.Write (" selected") %>>Sucata</option> 
  			</select>         
        </td>
      </tr>
      <tr>
        <td class="fundo1" width="90" height="65" align="center"><b>Observa��o</b></td>
        <td class="fundo3" width="510" height="65" colspan="3"><textarea name="estabilizador_observa" cols="57" rows="6" style="border-style: inset; border-width: 5"><% = rsbanco("estabilizador_observa")%></textarea></td>
      </tr>                
</table>
<input type="submit" value="&nbsp;&nbsp;Salvar&nbsp;&nbsp;" name="BTincluir">
</form>
<form action="consultas4.asp"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;">
<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2" color="#FFFFFF">OBS: Todos os Campos s�o obrigat�rios. Verifique as informa��es antes de salv�-las,</font></b></p>
<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2" color="#FFFFFF">&nbsp; pois as mesmas n�o poder�o ser apagadas pelo usu�rio solicitante.</font></b></p>
<!--#include file="rodape.asp"--></form>
</center></body></html>
<% Else
	Response.Redirect("admin.asp")
End If %>