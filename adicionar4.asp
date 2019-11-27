<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% If Session("LoggedIn") = True Then %>
<body><center><form method="GET" action="add6.asp">
<table border="1" width="600" height="102">
    <tr>
      <td class="fundo1" width="90" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="510" height="102" align="center">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Inventário de Informática</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Cadastro de Hardware</b></font></td>
    </tr>
</table>
<table border="1" width="600" height="1">
      <tr>
        <td class="fundo1" width="219" height="23" align="center"><b>Data</b></td>
        <td class="fundo3" width="24" height="23"><input readonly name="computador_data" type="text" value="<% Response.write Now %>" size="20" style="border-style: inset; border-width: 5; text-align:center"></td>
        <td class="fundo1" width="116" height="23" align="center"><b>Periférico</b></td>
        <td class="fundo3" width="372" height="23"><input name="computador_periferico" type="text" readOnly value="Computador" size="20" style="border-style: inset; border-width: 5; text-align:center"></td>
      </tr>     
      <tr>
        <td class="fundo1" width="219" height="23" align="center"><b>Seção</b></td>
        <td class="fundo3" width="194" height="23">
<% 	 Set rsbanco3 = Server.CreateObject("ADODB.Recordset")
		rsbanco3.Open "SELECT * FROM secao ORDER BY secao_nome ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select readOnly name="computador_secao">
				<option value="0" selected>Selecione</option>								    
		<%  Do While Not rsbanco3.EOF %>								    
				<option value="<% = rsbanco3("secao_nome") %>"><% = rsbanco3("secao_nome") %></option>
		<% 		rsbanco3.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco3.Close
		Set rsbanco3= Nothing %>
        </td>
		<td class="fundo1" width="116" height="23" align="center"><b>Esquadrão</b></td>
        <td class="fundo3" width="372" height="23">
<% If Session("Level") = 4 Then %>         
		<input name="computador_esquadrao" type="text" readOnly value="<% = Session("Esquadrao")%>" size="20" style="border-style: inset; border-width: 5; text-align:center">        
<% Else
 	Set rsbanco4 = Server.CreateObject("ADODB.Recordset")
		rsbanco4.Open "SELECT * FROM esquadrao ORDER BY esquadrao_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="computador_esquadrao">					    
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
        <td class="fundo1" width="219" height="23" align="center"><b>FCG</b></td>
        <td class="fundo3" width="194" height="23"><input type="text" value="0" name="computador_fcg" size="12" style="border-style: inset; border-width: 5; text-align:center"></td>
        <td class="fundo1" width="116" height="23" align="center"><b>Tipo</b></td>
        <td class="fundo3" width="372" valign="middle" height="23">
            <select size="1" name="computador_tipo">
  				<option value="0" selected>Selecione</option>            
  				<option>Cliente</option>
  				<option>Servidor</option> 
  				<option>Grande Porte</option>  				  				 				            
  			</select>         
        </td>
      </tr>
      <tr>
        <td class="fundo1" width="219" height="23" align="center"><b>Sist Oper</b></td>
        <td class="fundo3" width="194" height="23">
            <select size="1" name="computador_so">
  				<option value="0" selected>Selecione</option>
  				<option>Linux</option>  				           
  				<option>Windows 95</option>
  				<option>Windows 98</option> 
  				<option>Windows ME</option>
  				<option>Windows NT</option>  				 				  				 				            
  				<option>Windows XP</option>
  				<option>Windows 2000</option>
  				<option>Windows Vista</option>
  			</select>        
        </td>
        <td class="fundo1" width="116" height="23" align="center"><b>Qtd Proces</b></td>
        <td class="fundo3" width="372" valign="middle" height="23">
            <select size="1" name="computador_qp">
  				<option value="0" selected>Selecione</option>
  				<option>1</option>  				           
  				<option>2</option>
  				<option>3</option> 
  				<option>4</option>
  				<option>5</option>  				 				  				 				            
  				<option>6</option>
  				<option>7</option>
  				<option>8</option>
  			</select>         
        </td>
      </tr>
      <tr>
        <td class="fundo1" width="219" height="23" align="center"><b>Processador</b></td>
        <td class="fundo3" width="194" height="23"><input type="text" name="computador_procvelo" size="6" style="border-style: inset; border-width: 5">
            <select size="1" name="computador_procfreq">
  				<option value="0" selected>Selecione</option>
  				<option>MHz</option>  				           
  				<option>GHz</option>
  				<option>MIps</option> 
  			</select> 
        </td>
        <td class="fundo1" width="116" height="23" align="center"><b>Memória</b></td>
        <td class="fundo3" width="372" valign="middle" height="23"><input type="text" name="computador_memovelo" size="6" style="border-style: inset; border-width: 5">
            <select size="1" name="computador_memocapa">
  				<option value="0" selected>Selecione</option>
  				<option>MB</option>  				           
  				<option>GB</option>
  				<option>TB</option> 
  			</select>         
        </td>
      </tr>        
      <tr>
        <td class="fundo1" width="219" height="23" align="center"><b>Hard Disk</b></td>
        <td class="fundo3" width="194" height="23"><input type="text" name="computador_hdtama" size="6" style="border-style: inset; border-width: 5">
            <select size="1" name="computador_hdcapa">
  				<option value="0" selected>Selecione</option>
  				<option>MB</option>  				           
  				<option>GB</option>
  				<option>TB</option> 
  			</select> 
  	    </td>
        <td class="fundo1" width="116" height="23" align="center"><b>Situação</b></td>
        <td class="fundo3" width="372" valign="middle" height="23">
            <select size="1" name="computador_situacao">
  				<option value="0" selected>Selecione</option>
  				<option>Uso</option>  				           
  				<option>Manutenção</option>
  				<option>Sucata</option> 
  			</select>         
        </td>
      </tr>
      <tr>
        <td class="fundo1" width="214" height="65" align="center"><b>Observação</b></td>
        <td class="fundo3" width="387" height="65" colspan="3"><textarea name="computador_observa" cols="57" rows="6" style="border-style: inset; border-width: 5">Nada a relatar</textarea></td>
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