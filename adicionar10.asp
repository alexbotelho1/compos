<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% If Session("LoggedIn") = True Then %>
<body><center>
<table border="1" width="700" height="102">
    <tr>
      <td class="fundo1" width="90" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="610" height="102" align="center">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Inventário de Informática</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Cadastrar Diversos</b></font></td>
    </tr>
</table>
<% If Session("Level") = 1 Then %>
<table border="0" cellpadding="0" cellspacing="0" width="700" height="35">
	<tr>
		<td width="700" height="35"><form method="GET" action="alt10.asp">
			<% 	Set rsbanco10 = Server.CreateObject("ADODB.Recordset")
					rsbanco10.Open "SELECT * FROM config ORDER BY config_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic %>	
			<table border="1" height="35" width="700">    
				<tr>
        			<td class="fundo1" width="291" height="35" align="center"><b>Página &quot;Index&quot; de Manutenção</b></td>
	        		<td class="fundo3" width="312" height="35" align="center">
	        		<input type="radio" value="0" name="config_manu"<% If rsbanco10("config_manu") = 0 then Response.Write (" checked ") %>> Desativar
	        		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	        		<input type="radio" value="1" name="config_manu"<% If rsbanco10("config_manu") = 1 then Response.Write (" checked ") %>> Ativar</td>
   					<td class="fundo5" width="97" height="35" align="center"><input type="submit" value="Salvar">
				<% 	rsbanco10.Close
					Set rsbanco10 = Nothing %>  					
   					</td>
    			</tr>
			</table>        
		</td></form>
	</tr>                
</table>
<% End If %> 
<form  action="admin.asp"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;"><!--#include file="rodape.asp"--></form></center></body></html>
<% Else
	Response.Redirect("admin.asp")
End If %>