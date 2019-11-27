<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% If Session("LoggedIn") = True Then
	codigo = request.querystring("codigo")

	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from impressora WHERE impressora_codigo = "&codigo&"",banco,AdOpenKeySet,AdLockOptimistic
%>
<body><center><form method="GET" action="alt4.asp">
<table border="1" width="600" height="102">
    <tr>
      <td class="fundo1" width="86" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="514" height="102" align="center">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Inventário de Informática</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Alterar Impressora</b></font></td>
    </tr>
</table>
<table border="1" width="599" height="1">
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Data</b></td>
        <td class="fundo3" width="210" height="23"><input name="impressora_data" type="text" value="<% Response.write Now %>" size="20" style="border-style: inset; border-width: 5; text-align:center"></td>
        <td class="fundo1" width="90" height="23" align="center"><b>FCG</b></td>
        <td class="fundo3" width="210" height="23"><input name="impressora_fcg" type="text" value="<% = rsbanco("impressora_fcg") %>" size="12" style="border-style: inset; border-width: 5"></td>
      </tr>
			<input type=hidden name=impressora_codigo value="<%=rsbanco("impressora_codigo")%>">      
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Seção</b></td>
        <td class="fundo3" width="210" height="23">
	<% 	Set rsbanco3 = Server.CreateObject("ADODB.Recordset")
		rsbanco3.Open "SELECT * FROM secao ORDER BY secao_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="impressora_secao">				    
		<%  Do While Not rsbanco3.EOF %>								    
				<option value="<% = rsbanco3("secao_nome") %>"<% If rsbanco("impressora_secao") = rsbanco3("secao_nome") Then Response.Write (" selected") %>><% = rsbanco3("secao_nome") %></option>
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
			<select name="impressora_esquadrao">				    
		<%  Do While Not rsbanco4.EOF %>								    
				<option value="<% = rsbanco4("esquadrao_nome") %>"<% If rsbanco("impressora_esquadrao") = rsbanco4("esquadrao_nome") Then Response.Write (" selected") %>><% = rsbanco4("esquadrao_nome") %></option>
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
		rsbanco2.Open "SELECT * FROM marcaimp ORDER BY marcaimp_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="impressora_marca">	 
		<%  Do While Not rsbanco2.EOF %>								    
				<option value="<% = rsbanco2("marcaimp_nome") %>"<% If rsbanco("impressora_marca") = rsbanco2("marcaimp_nome") Then Response.Write (" selected") %>><% = rsbanco2("marcaimp_nome") %></option>
		<% 		rsbanco2.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco2.Close
		Set rsbanco2 = Nothing %>   
        </td>      
        <td class="fundo1" width="90" height="23" align="center"><b>Modelo</b></td>
        <td class="fundo3" width="210" height="23">
	<% 	Set rsbanco1 = Server.CreateObject("ADODB.Recordset")
		rsbanco1.Open "SELECT * FROM modimp ORDER BY modimp_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="impressora_modelo">	 
		<%  Do While Not rsbanco1.EOF %>								    
				<option value="<% = rsbanco1("modimp_nome") %>"<% If rsbanco("impressora_modelo") = rsbanco1("modimp_nome") Then Response.Write (" selected") %>><% = rsbanco1("modimp_nome") %></option>
		<% 		rsbanco1.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco1.Close
		Set rsbanco1 = Nothing %>        
		</td>
      </tr>
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Impressão</b></td>
        <td class="fundo3" width="210" height="23">
            <select size="1" name="impressora_impressao">
  				<option<% If rsbanco("impressora_impressao") = "Laser" Then Response.Write (" selected") %>>
                Laser</option><option<% If rsbanco("impressora_impressao") = "Tinta" Then Response.Write (" selected") %>>Tinta</option>
  			</select>        
        </td>
        <td class="fundo1" width="90" height="23" align="center"><b>Cor</b></td>
        <td class="fundo3" width="210" height="23">
            <select size="1" name="impressora_cor">
  				<option<% If rsbanco("impressora_cor") = "Preto" Then Response.Write (" selected") %>>
                Preto</option><option<% If rsbanco("impressora_cor") = "Color" Then Response.Write (" selected") %>>Color</option>
  			</select>         
        </td>
      </tr>
      <tr>
        <td class="fundo1" width="600" height="23" colspan="4" align="center"><b>Modelos dos Cartuchos</b></td>
      </tr>      
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Colorido</b></td>
        <td class="fundo3" width="210" height="23"><input type="text" value="<% = rsbanco("impressora_colorido")%>" name="impressora_colorido" size="6" style="border-style: inset; border-width: 5"></td>
        <td class="fundo1" width="90" height="23" align="center"><b>Preto</b></td>
        <td class="fundo3" width="210" height="23"><input type="text" value="<% = rsbanco("impressora_preto")%>" name="impressora_preto" size="6" style="border-style: inset; border-width: 5"></td>
      </tr>
      <tr>
        <td class="fundo1" width="300" height="23" colspan="2" align="center"><b>Modelo do Toner</b></td>
        <td class="fundo3" width="300" height="23" colspan="2"><input type="text" value="<% = rsbanco("impressora_toner")%>" name="impressora_toner" size="6" style="border-style: inset; border-width: 5"></td>
      </tr>        
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Consumo</b></td>
        <td class="fundo3" width="210" height="23">
            <select size="1" name="impressora_consumo">
  				<option value="0.5"<% If rsbanco("impressora_consumo") = "0.5" Then Response.Write (" selected") %>>0,5 = 1Q/2M</option>  				           
  				<option value="1"<% If rsbanco("impressora_consumo") = "1.0" Then Response.Write (" selected") %>>1,0 = 1Q/1M</option>
  				<option value="1.5"<% If rsbanco("impressora_consumo") = "1.5" Then Response.Write (" selected") %>>1,5 = 3Q/2M</option>
  				<option value="2"<% If rsbanco("impressora_consumo") = "2.0" Then Response.Write (" selected") %>>2,0 = 2Q/1M</option>
 				<option value="2.5"<% If rsbanco("impressora_consumo") = "2.5" Then Response.Write (" selected") %>>2,5 = 5Q/2M</option>
  				<option value="3"<% If rsbanco("impressora_consumo") = "3.0" Then Response.Write (" selected") %>>3,0 = 3Q/1M</option>
 				<option value="3.5"<% If rsbanco("impressora_consumo") = "3.5" Then Response.Write (" selected") %>>3,5 = 7Q/2M</option>
  				<option value="4"<% If rsbanco("impressora_consumo") = "4.0" Then Response.Write (" selected") %>>4,0 = 4Q/1M</option>  				
  			</select>&nbsp;<b><font size="2" color="#000080">QTD (</font><font size="2" color="#FF0000">Q</font><font size="2" color="#000080">)</font><font size="4" color="#FF0000">/</font><font size="2" color="#000080">(</font><font size="2" color="#FF0000">M</font><font size="2" color="#000080">)MÊS</font></b></td>
        <td class="fundo1" width="90" height="23" align="center"><b>Situação</b></td>
        <td class="fundo3" width="210" height="23">
            <select size="1" name="impressora_situacao">
  				<option value="Uso"<% If rsbanco("impressora_situacao") = "Uso" Then Response.Write (" selected") %>>Uso</option>
                <option value="Manutenção"<% If rsbanco("impressora_situacao") = "Manutenção" Then Response.Write (" selected") %>>Manutenção</option>
                <option value="Sucata"<% If rsbanco("impressora_situacao") = "Sucata" Then Response.Write (" selected") %>>Sucata</option> 
  			</select>         
        </td>
      </tr>
      <tr>
        <td class="fundo1" width="90" height="65" align="center"><b>Observação</b></td>
        <td class="fundo3" width="510" height="65" colspan="3"><textarea name="impressora_observa" cols="57" rows="6" style="border-style: inset; border-width: 5"><% = rsbanco("impressora_observa")%></textarea></td>
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