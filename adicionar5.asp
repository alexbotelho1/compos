<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% If Session("LoggedIn") = True Then %>
<body><center><form method="GET" action="add7.asp">
<table border="1" width="600" height="102">
    <tr>
      <td class="fundo1" width="86" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="514" height="102" align="center">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Invent�rio de Inform�tica</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Adicionar Impressora</b></font></td>
    </tr>
</table>
<table border="1" width="599" height="1">
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Data</b></td>
        <td class="fundo3" width="210" height="23"><input readonly name="impressora_data" type="text" value="<% Response.write Now %>" size="20" style="border-style: inset; border-width: 5; text-align:center"></td>
        <td class="fundo1" width="90" height="23" align="center"><b>FCG</b></td>
        <td class="fundo3" width="210" height="23"><input type="text" value="0" name="impressora_fcg" size="12" style="border-style: inset; border-width: 5; text-align:center"></td>
      </tr>     
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Se��o</b></td>
        <td class="fundo3" width="210" height="23">
	<% 	Set rsbanco3 = Server.CreateObject("ADODB.Recordset")
		rsbanco3.Open "SELECT * FROM secao ORDER BY secao_nome ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="impressora_secao">
				<option value="0" selected>Selecione</option>					    
		<%  Do While Not rsbanco3.EOF %>								    
				<option value="<% = rsbanco3("secao_nome") %>"><% = rsbanco3("secao_nome") %></option>
		<% 		rsbanco3.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco3.Close
		Set rsbanco3= Nothing %>
        </td>
		<td class="fundo1" width="90" height="23" align="center"><b>Esquadr�o</b></td>
        <td class="fundo3" width="210" height="23">
<% If Session("Level") = 4 Then %>         
		<input name="impressora_esquadrao" type="text" readOnly value="<% = Session("Esquadrao")%>" size="20" style="border-style: inset; border-width: 5; text-align:center">        
<% Else
 	Set rsbanco4 = Server.CreateObject("ADODB.Recordset")
		rsbanco4.Open "SELECT * FROM esquadrao ORDER BY esquadrao_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="impressora_esquadrao">					    
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
		rsbanco2.Open "SELECT * FROM marcaimp ORDER BY marcaimp_nome ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="impressora_marca">	
				<option value="0" selected>Selecione</option>				    
		<%  Do While Not rsbanco2.EOF %>								    
				<option value="<% = rsbanco2("marcaimp_nome") %>"><% = rsbanco2("marcaimp_nome") %></option>
		<% 		rsbanco2.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco2.Close
		Set rsbanco2 = Nothing %>   
        </td>      
        <td class="fundo1" width="90" height="23" align="center"><b>Modelo</b></td>
        <td class="fundo3" width="210" height="23">
	<% 	Set rsbanco1 = Server.CreateObject("ADODB.Recordset")
		rsbanco1.Open "SELECT * FROM modimp ORDER BY modimp_nome ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="impressora_modelo">	
				<option value="0" selected>Selecione</option>				    
		<%  Do While Not rsbanco1.EOF %>								    
				<option value="<% = rsbanco1("modimp_nome") %>"><% = rsbanco1("modimp_nome") %></option>
		<% 		rsbanco1.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco1.Close
		Set rsbanco1 = Nothing %>        
		</td>
      </tr>
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Impress�o</b></td>
        <td class="fundo3" width="210" height="23">
            <select size="1" name="impressora_impressao">
  				<option value="0" selected>Selecione</option>
  				<option>Laser</option>  				           
  				<option>Tinta</option>
  				<option>Fita</option>
  			</select>        
        </td>
        <td class="fundo1" width="90" height="23" align="center"><b>Cor</b></td>
        <td class="fundo3" width="210" valign="middle" height="23">
            <select size="1" name="impressora_cor">
  				<option value="0" selected>Selecione</option>
  				<option>Preto</option>  				           
  				<option>Color</option>
  			</select>         
        </td>
      </tr>
      <tr>
        <td class="fundo1" width="600" height="23" colspan="4" align="center"><b>Modelos dos Cartuchos</b></td>
      </tr>      
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Colorido</b></td>
        <td class="fundo3" width="210" height="23"><input type="text" value="0" name="impressora_colorido" size="6" style="border-style: inset; border-width: 5"></td>
        <td class="fundo1" width="90" height="23" align="center"><b>Preto</b></td>
        <td class="fundo3" width="210" valign="middle" height="23"><input type="text" value="0" name="impressora_preto" size="6" style="border-style: inset; border-width: 5"></td>
      </tr>
      <tr>
        <td class="fundo1" width="300" height="23" colspan="2" align="center"><b>Modelo do Toner</b></td>
        <td class="fundo3" width="300" height="23" colspan="2"><input type="text" value="0" name="impressora_toner" size="6" style="border-style: inset; border-width: 5"></td>
      </tr>        
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Consumo</b></td>
        <td class="fundo3" width="210" height="23">
            <select size="1" name="impressora_consumo">
  				<option value="0" selected>Selecione</option>
  				<option value="0.5">0,5 = 1Q/2M</option>  				           
  				<option value="1">1,0 = 1Q/1M</option>
  				<option value="1.5">1,5 = 3Q/2M</option>
  				<option value="2">2,0 = 2Q/1M</option>
 				<option value="2.5">2,5 = 5Q/2M</option>
  				<option value="3">3,0 = 3Q/1M</option>
 				<option value="3.5">3,5 = 7Q/2M</option>
  				<option value="4">4,0 = 4Q/1M</option>  				
  			</select>&nbsp;<b><font size="2" color="#000080">QTD (</font><font size="2" color="#FF0000">Q</font><font size="2" color="#000080">)</font><font size="4" color="#FF0000">/</font><font size="2" color="#000080">(</font><font size="2" color="#FF0000">M</font><font size="2" color="#000080">)M�S</font></b></td>
        <td class="fundo1" width="90" height="23" align="center"><b>Situa��o</b></td>
        <td class="fundo3" width="210" valign="middle" height="23">
            <select size="1" name="impressora_situacao">
  				<option value="0" selected>Selecione</option>
  				<option>Uso</option>  				           
  				<option>Manuten��o</option>
  				<option>Sucata</option> 
  			</select>         
        </td>
      </tr>
      <tr>
        <td class="fundo1" width="90" height="65" align="center"><b>Observa��o</b></td>
        <td class="fundo3" width="510" height="65" colspan="3"><textarea name="impressora_observa" cols="57" rows="6" style="border-style: inset; border-width: 5">Nada a relatar</textarea></td>
      </tr>                
    </table>
<input type="submit" value="Salvar" name="BTincluir">&nbsp;&nbsp;&nbsp;&nbsp;<input type="reset" value="Limpar" name="BTlimpar"></form>
<% If Session("Level") = 4 Then %>
<form action="admin.asp"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;">
<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2" color="#FFFFFF">OBS: Todos os Campos s�o obrigat�rios. Verifique as informa��es antes de salv�-las,</font></b></p>
<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2" color="#FFFFFF">&nbsp; pois as mesmas n�o poder�o ser apagadas pelo usu�rio solicitante.</font></b></p>
<!--#include file="rodape.asp"--></form>
<% Else %>
<form action="cad_hard.asp"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;">
<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2" color="#FFFFFF">OBS: Todos os Campos s�o obrigat�rios. Verifique as informa��es antes de salv�-las,</font></b></p>
<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2" color="#FFFFFF">&nbsp; pois as mesmas n�o poder�o ser apagadas pelo usu�rio solicitante.</font></b></p>
<!--#include file="rodape.asp"--></form>
<% End If %>
</center></body></html>
<% Else
	Response.Redirect("admin.asp")
End If %>