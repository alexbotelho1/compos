<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% If Session("LoggedIn") = True Then %>
<body><center><form method="GET" action="add8.asp">
<table border="1" width="600" height="102">
    <tr>
      <td class="fundo1" width="86" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="514" height="102" align="center">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Inventário de Informática</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Adicionar NoBreak</b></font></td>
    </tr>
</table>
<table border="1" width="599" height="1">
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Data</b></td>
        <td class="fundo3" width="210" height="23"><input readonly name="nobreak_data" type="text" value="<% Response.write Now %>" size="20" style="border-style: inset; border-width: 5; text-align:center"></td>
        <td class="fundo1" width="90" height="23" align="center"><b>FCG</b></td>
        <td class="fundo3" width="210" height="23"><input type="text" name="nobreak_fcg" size="12" style="border-style: inset; border-width: 5"></td>
      </tr>     
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Seção</b></td>
        <td class="fundo3" width="210" height="23">
	<% 	Set rsbanco3 = Server.CreateObject("ADODB.Recordset")
		rsbanco3.Open "SELECT * FROM secao ORDER BY secao_nome ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="nobreak_secao">
				<option value="0" selected>Selecione</option>					    
		<%  Do While Not rsbanco3.EOF %>								    
				<option value="<% = rsbanco3("secao_nome") %>"><% = rsbanco3("secao_nome") %></option>
		<% 		rsbanco3.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco3.Close
		Set rsbanco3= Nothing %>
        </td>
		<td class="fundo1" width="90" height="23" align="center"><b>Esquadrão</b></td>
        <td class="fundo3" width="210" height="23">
<% If Session("Level") = 4 Then %>         
		<input name="nobreak_esquadrao" type="text" readOnly value="<% = Session("Esquadrao")%>" size="20" style="border-style: inset; border-width: 5; text-align:center">        
<% Else
 	Set rsbanco4 = Server.CreateObject("ADODB.Recordset")
		rsbanco4.Open "SELECT * FROM esquadrao ORDER BY esquadrao_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="nobreak_esquadrao">					    
		<%  Do While Not rsbanco4.EOF %>								    
				<option value="<% = rsbanco4("esquadrao_nome") %>"<% If rsbanco4("esquadrao_nome") = Session("Esquadrao") Then Response.Write (" selected") %>><% = rsbanco4("esquadrao_nome") %></option>
		<% 		rsbanco4.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco4.Close
		Set rsbanco4= Nothing
End If %>
        </td>
      </tr>      
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Marca</b></td>
        <td class="fundo3" width="210" valign="middle" height="23">
	<% 	Set rsbanco2 = Server.CreateObject("ADODB.Recordset")
		rsbanco2.Open "SELECT * FROM marcanb ORDER BY marcanb_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="nobreak_marca">	
				<option value="0" selected>Selecione</option>				    
		<%  Do While Not rsbanco2.EOF %>								    
				<option value="<% = rsbanco2("marcanb_nome") %>"><% = rsbanco2("marcanb_nome") %></option>
		<% 		rsbanco2.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco2.Close
		Set rsbanco2 = Nothing %>   
        </td>      
        <td class="fundo1" width="90" height="23" align="center"><b>Potência</b></td>
        <td class="fundo3" width="210" height="23">	  
	<% 	Set rsbanco1 = Server.CreateObject("ADODB.Recordset")
		rsbanco1.Open "SELECT * FROM nobreakpt ORDER BY nobreakpt_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="nobreak_potencia">	
				<option value="0" selected>Selecione</option>				    
		<%  Do While Not rsbanco1.EOF %>								    
				<option value="<% = rsbanco1("nobreakpt_nome") %>"><% = rsbanco1("nobreakpt_nome") %></option>
		<% 		rsbanco1.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco1.Close
		Set rsbanco1 = Nothing %> <b><font size="2" color="#0000FF">KVa</font></b></td>
      </tr>   
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Qtd Saída</b></td>
        <td class="fundo3" width="210" height="23">
            <select size="1" name="nobreak_saida">
  				<option value="0" selected>Selecione</option>
  				<option>1</option>  				           
  				<option>2</option>
  				<option>3</option>
  				<option>4</option>
  				<option>5</option>
  				<option>6</option>
  				<option>7</option>  				  				  				  				  								
  			</select>
  		</td>
        <td class="fundo1" width="90" height="23" align="center"><b>Situação</b></td>
        <td class="fundo3" width="210" valign="middle" height="23">
            <select size="1" name="nobreak_situacao">
  				<option value="0" selected>Selecione</option>
  				<option>Uso</option>  				           
  				<option>Manutenção</option>
  				<option>Sucata</option> 
  			</select>         
        </td>
      </tr>
      <tr>
        <td class="fundo1" width="90" height="65" align="center"><b>Observação</b></td>
        <td class="fundo3" width="510" height="65" colspan="3"><textarea name="nobreak_observa" cols="57" rows="6" style="border-style: inset; border-width: 5">Nada a relatar</textarea></td>
      </tr>                
</table>
<input type="submit" value="Salvar" name="BTincluir">&nbsp;&nbsp;&nbsp;&nbsp;<input type="reset" value="Limpar" name="BTlimpar"></form>
<% If Session("Level") = 4 Then %>
<form action="admin.asp"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;">
<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2" color="#FFFFFF">OBS: Todos os Campos são obrigatórios. Verifique as informações antes de salvá-las,</font></b></p>
<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2" color="#FFFFFF">&nbsp; pois as mesmas não poderão ser apagadas pelo usuário solicitante.</font></b></p>
<!--#include file="rodape.asp"--></form>
<% Else %>
<form action="cad_hard.asp"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;">
<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2" color="#FFFFFF">OBS: Todos os Campos são obrigatórios. Verifique as informações antes de salvá-las,</font></b></p>
<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2" color="#FFFFFF">&nbsp; pois as mesmas não poderão ser apagadas pelo usuário solicitante.</font></b></p>
<!--#include file="rodape.asp"--></form>
<% End If %>
</center></body></html>
<% Else
	Response.Redirect("admin.asp")
End If %>