<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% If Session("LoggedIn") = True Then	
	codigo = request.querystring("codigo")

	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from impressora WHERE impressora_codigo = "&codigo&"",banco,AdOpenKeySet,AdLockOptimistic
%>
<body><center><form action="exc4.asp">
<table border="1" width="600" height="102">
    <tr>
      <td class="fundo1" width="86" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="514" height="102" align="center">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Inventário de Informática</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Excluir Impressora</b></font></td>
    </tr>
</table>
<table border="1" width="599" height="1">
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Data</b></td>
        <td class="fundo3" width="210" height="23"><% = rsbanco("impressora_data") %></td>
        <td class="fundo1" width="90" height="23" align="center"><b>FCG</b></td>
        <td class="fundo3" width="210" height="23"><% = rsbanco("impressora_fcg") %></td>
      </tr>
<input type=hidden name=impressora_codigo value="<%=rsbanco("impressora_codigo")%>">
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Seção</b></td>
        <td class="fundo3" width="210" height="23"><% = rsbanco("impressora_secao") %></td>
		<td class="fundo1" width="90" height="23" align="center"><b>Esquadrão</b></td>
        <td class="fundo3" width="210" height="23"><% = rsbanco("impressora_esquadrao") %></td>
      </tr>      
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Marca</b></td>
        <td class="fundo3" width="210" valign="middle" height="23"><% = rsbanco("impressora_marca") %></td>      
        <td class="fundo1" width="90" height="23" align="center"><b>Modelo</b></td>
        <td class="fundo3" width="210" height="23"><% = rsbanco("impressora_modelo") %></td>
      </tr>
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Impressão</b></td>
        <td class="fundo3" width="210" height="23"><% = rsbanco("impressora_impressao") %></td>
        <td class="fundo1" width="90" height="23" align="center"><b>Cor</b></td>
        <td class="fundo3" width="210" valign="middle" height="23"><% = rsbanco("impressora_cor") %></td>
      </tr>
      <tr>
        <td class="fundo1" width="600" height="23" colspan="4" align="center"><b>Modelos dos Cartuchos</b></td>
      </tr>      
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Colorido</b></td>
        <td class="fundo3" width="210" height="23"><% = rsbanco("impressora_colorido") %></td>
        <td class="fundo1" width="90" height="23" align="center"><b>Preto</b></td>
        <td class="fundo3" width="210" valign="middle" height="23"><% = rsbanco("impressora_preto") %></td>
      </tr>
      <tr>
        <td class="fundo1" width="300" height="23" colspan="2" align="center"><b>Modelo do Toner</b></td>
        <td class="fundo3" width="300" height="23" colspan="2"><% = rsbanco("impressora_toner") %></td>
      </tr>        
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Consumo</b></td>
        <td class="fundo3" width="210" height="23"><% = rsbanco("impressora_consumo") %>&nbsp;<b><font size="2" color="#000080">QTD (</font><font size="2" color="#FF0000">Q</font><font size="2" color="#000080">)</font><font size="4" color="#FF0000">/</font><font size="2" color="#000080">(</font><font size="2" color="#FF0000">M</font><font size="2" color="#000080">)MÊS</font></b></td>
        <td class="fundo1" width="90" height="23" align="center"><b>Situação</b></td>
        <td class="fundo3" width="210" valign="middle" height="23"><% = rsbanco("impressora_situacao") %></td>
      </tr>
      <tr>
        <td class="fundo1" width="90" height="65" align="center"><b>Observação</b></td>
        <td class="fundo3" width="510" height="65" colspan="3"><textarea cols="57" rows="6" style="border-style: inset; border-width: 5"><% = rsbanco("impressora_observa") %></textarea></td>
      </tr>                
</table>
<input type="submit" value="&nbsp;&nbsp;Excluir&nbsp;&nbsp;" name="BTincluir"></form>
<form action="consultas4.asp"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;">
<!--#include file="rodape.asp"--></form></body></html>
<% Else
	Response.Redirect("admin.asp")
End If %>