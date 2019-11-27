<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% If Session("LoggedIn") = True Then	
	codigo = request.querystring("codigo")

	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from estabilizador WHERE estabilizador_codigo = "&codigo&"",banco,AdOpenKeySet,AdLockOptimistic
%>
<body><center><form action="exc6.asp">
<table border="1" width="600" height="102">
    <tr>
      <td class="fundo1" width="86" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="514" height="102" align="center">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Inventário de Informática</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Excluir Estabilizador</b></font></td>
    </tr>
</table>
<table border="1" width="599" height="1">
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Data</b></td>
        <td class="fundo3" width="210" height="23"><% = rsbanco("estabilizador_data") %></td>
        <td class="fundo1" width="90" height="23" align="center"><b>FCG</b></td>
        <td class="fundo3" width="210" height="23"><% = rsbanco("estabilizador_fcg") %></td>
      </tr>
<input type=hidden name=estabilizador_codigo value="<%=rsbanco("estabilizador_codigo")%>">
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Seção</b></td>
        <td class="fundo3" width="210" height="23"><% = rsbanco("estabilizador_secao") %></td>
		<td class="fundo1" width="90" height="23" align="center"><b>Esquadrão</b></td>
        <td class="fundo3" width="210" height="23"><% = rsbanco("estabilizador_esquadrao") %></td>
      </tr>      
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Marca</b></td>
        <td class="fundo3" width="210" valign="middle" height="23"><% = rsbanco("estabilizador_marca") %></td>      
		<td class="fundo1" width="90" height="23" align="center"><b>Situação</b></td>
        <td class="fundo3" width="210" valign="middle" height="23"><% = rsbanco("estabilizador_situacao") %></td>
      </tr>
      <tr>
        <td class="fundo1" width="90" height="65" align="center"><b>Observação</b></td>
        <td class="fundo3" width="510" height="65" colspan="3"><% = rsbanco("estabilizador_observa") %></td>
      </tr>                
</table>
<input type="submit" value="&nbsp;&nbsp;Excluir&nbsp;&nbsp;" name="BTincluir"></form>
<form action="consultas4.asp"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;">
<!--#include file="rodape.asp"--></form></body></html>
<% Else
	Response.Redirect("admin.asp")
End If %>