<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% If Session("LoggedIn") = True Then
	codigo = request.querystring("codigo")

	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from computador WHERE computador_codigo = "&codigo&"",banco,AdOpenKeySet,AdLockOptimistic
%>
<body><center><form method="GET" action="alt3.asp">
<table border="1" width="600" height="102">
    <tr>
      <td class="fundo1" width="90" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td width="510" height="102" align="center" class="fundo2">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Inventário de Informática</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Alterar Cadastro de Hardware</b></font></td>
    </tr>
</table>
<table border="1" width="599" height="1">
      <tr>
        <td class="fundo1" width="226" height="23" align="center"><b>Data</b></td>
        <td class="fundo3" width="73" height="23"><input name="computador_data" type="text" value="<%=rsbanco("computador_data")%>" size="20" style="border-style: inset; border-width: 5; text-align:center"></td>
        <td class="fundo1" width="94" height="23" align="center"><b>Periférico</b></td>
        <input type=hidden name=computador_codigo value="<%=rsbanco("computador_codigo")%>">
        <td class="fundo3" width="337" height="23">
	<% 	Set rsbanco6 = Server.CreateObject("ADODB.Recordset")
		rsbanco6.Open "SELECT * FROM hardware ORDER BY hardware_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="computador_periferico">					    
		<%  Do While Not rsbanco6.EOF %>								    
				<option value="<% = rsbanco6("hardware_nome") %>"<% If rsbanco("computador_periferico") = rsbanco6("hardware_nome") Then Response.Write (" selected") %>><% = rsbanco6("hardware_nome") %></option>
		<% 		rsbanco6.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco6.Close
		Set rsbanco6 = Nothing %>        
        </td>
      </tr>     
      <tr>
        <td class="fundo1" width="226" height="23" align="center"><b>Seção</b></td>
        <td class="fundo3" width="243" height="23">
	<% 	Set rsbanco3 = Server.CreateObject("ADODB.Recordset")
		rsbanco3.Open "SELECT * FROM secao ORDER BY secao_nome ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="computador_secao">					    
		<%  Do While Not rsbanco3.EOF %>								    
				<option value="<% = rsbanco3("secao_nome") %>"<% If rsbanco("computador_secao") = rsbanco3("secao_nome") Then Response.Write (" selected") %>><% = rsbanco3("secao_nome") %></option>
		<% 		rsbanco3.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco3.Close
		Set rsbanco3= Nothing %>        
  		</td>
		<td class="fundo1" width="94" height="23" align="center"><b>Esquadrão</b></td>
        <td class="fundo3" width="337" height="23">
	<% 	Set rsbanco4 = Server.CreateObject("ADODB.Recordset")
		rsbanco4.Open "SELECT * FROM esquadrao ORDER BY esquadrao_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="computador_esquadrao">					    
		<%  Do While Not rsbanco4.EOF %>								    
				<option value="<% = rsbanco4("esquadrao_nome") %>"<% If rsbanco("computador_esquadrao") = rsbanco4("esquadrao_nome") Then Response.Write (" selected") %>><% = rsbanco4("esquadrao_nome") %></option>
		<% 		rsbanco4.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco4.Close
		Set rsbanco4= Nothing %>       
		</td>
      </tr>      
      <tr>
        <td class="fundo1" width="226" height="23" align="center"><b>FCG</b></td>
        <td class="fundo3" width="243" height="23"><input name="computador_fcg" type="text" value="<%=rsbanco("computador_fcg")%>" size="12" style="border-style: inset; border-width: 5; text-align:center"></td>
        <td class="fundo1" width="94" height="23" align="center"><b>Tipo</b></td>
        <td class="fundo3" width="337" height="23">
            <select size="1" name="computador_tipo">  
  				<option<% If rsbanco("computador_tipo") = "Cliente" Then Response.Write (" selected") %>>
                Cliente</option><option<% If rsbanco("computador_tipo") = "Servidor" Then Response.Write (" selected") %>>Servidor</option><option<% If rsbanco("computador_tipo") = "Grande Porte" Then Response.Write (" selected") %>>Grande Porte</option>  				  				 				            
  			</select>        
	  	</td>
	  </tr>
      <tr>
        <td class="fundo1" width="226" height="23" align="center"><b>Sist Oper</b></td>
        <td class="fundo3" width="243" height="23">
            <select size="1" name="computador_so">
  				<option<% If rsbanco("computador_so") = "Linux" Then Response.Write (" selected") %>>
                Linux</option><option<% If rsbanco("computador_so") = "Windows 95" Then Response.Write (" selected") %>>Windows 95</option><option<% If rsbanco("computador_so") = "Windows 98" Then Response.Write (" selected") %>>Windows 98</option><option<% If rsbanco("computador_so") = "Windows ME" Then Response.Write (" selected") %>>Windows ME</option><option<% If rsbanco("computador_so") = "Windows NT" Then Response.Write (" selected") %>>Windows NT</option><option<% If rsbanco("computador_so") = "Windows XP" Then Response.Write (" selected") %>>Windows XP</option><option<% If rsbanco("computador_so") = "Windows 2000" Then Response.Write (" selected") %>>Windows 2000</option
  				><option<% If rsbanco("computador_so") = "Windows Vista" Then Response.Write (" selected") %>>Windows Vista</option>
  			</select>        
        </td>
        <td class="fundo1" width="94" height="23" align="center"><b>Qtd Proces</b></td>
        <td class="fundo3" width="337" height="23">
            <select size="1" name="computador_qp">
  				<option<% If rsbanco("computador_qp") = "1" Then Response.Write (" selected") %>>
                1</option><option<% If rsbanco("computador_qp") = "2" Then Response.Write (" selected") %>>2</option><option<% If rsbanco("computador_qp") = "3" Then Response.Write (" selected") %>>3</option><option<% If rsbanco("computador_qp") = "4" Then Response.Write (" selected") %>>4</option><option<% If rsbanco("computador_qp") = "5" Then Response.Write (" selected") %>>5</option><option<% If rsbanco("computador_qp") = "6" Then Response.Write (" selected") %>>6</option><option<% If rsbanco("computador_qp") = "7" Then Response.Write (" selected") %>>7</option><option<% If rsbanco("computador_qp") = "8" Then Response.Write (" selected") %>>8</option>
  			</select> 
		 </td>
      </tr>
      <tr>
        <td class="fundo1" width="226" height="23" align="center"><b>Processador</b></td>
        <td class="fundo3" width="243" height="23"><input name="computador_procvelo" type="text" value="<%=rsbanco("computador_procvelo")%>" size="6" style="border-style: inset; border-width: 5; text-align:center">
            <select size="1" name="computador_procfreq">
  				<option<% If rsbanco("computador_procfreq") = "MHz" Then Response.Write (" selected") %>>
                MHz</option><option<% If rsbanco("computador_procfreq") = "GHz" Then Response.Write (" selected") %>>GHz</option><option<% If rsbanco("computador_procfreq") = "MIps" Then Response.Write (" selected") %>>MIps</option> 
  			</select>
        </td>
        <td class="fundo1" width="94" height="23" align="center"><b>Memória</b></td>
        <td class="fundo3" width="337" height="23"><input name="computador_memovelo" type="text" value="<%=rsbanco("computador_memovelo")%>" size="6" style="border-style: inset; border-width: 5; text-align:center">
            <select size="1" name="computador_memocapa">
  				<option<% If rsbanco("computador_memocapa") = "MB" Then Response.Write (" selected") %>>
                MB</option><option<% If rsbanco("computador_memocapa") = "GB" Then Response.Write (" selected") %>>GB</option><option<% If rsbanco("computador_memocapa") = "TB" Then Response.Write (" selected") %>>TB</option> 
  			</select> 
        </td>
      </tr>        
      <tr>
        <td class="fundo1" width="226" height="23" align="center"><b>Hard Disk</b></td>
        <td class="fundo3" width="243" height="23"><input name="computador_hdtama" type="text" value="<%=rsbanco("computador_hdtama")%>" size="6" style="border-style: inset; border-width: 5; text-align:center">
            <select size="1" name="computador_hdcapa">
  				<option<% If rsbanco("computador_hdcapa") = "MB" Then Response.Write (" selected") %>>
                MB</option><option<% If rsbanco("computador_hdcapa") = "GB" Then Response.Write (" selected") %>>GB</option><option<% If rsbanco("computador_hdcapa") = "TB" Then Response.Write (" selected") %>>TB</option> 
  			</select> 
        </td>
        <td class="fundo1" width="94" height="23" align="center"><b>Situação</b></td>
        <td class="fundo3" width="337" height="23">
            <select size="1" name="computador_situacao">
  				<option<% If rsbanco("computador_situacao") = "Uso" Then Response.Write (" selected") %>>
                Uso</option><option<% If rsbanco("computador_situacao") = "Manutenção" Then Response.Write (" selected") %>>Manutenção</option><option<% If rsbanco("computador_situacao") = "Sucata" Then Response.Write (" selected") %>>Sucata</option> 
  			</select>        
      	</td>
      </tr>
      <tr>
        <td class="fundo1" width="221" height="65" align="center"><b>Observação</b></td>
        <td class="fundo3" width="379" height="65" colspan="3"><textarea name="computador_observa" cols="57" rows="6" style="border-style: inset; border-width: 5"><%=rsbanco("computador_observa")%></textarea></td>
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