<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!-- #include file="config.asp" -->
<!--#include file="styles.asp"-->
<%	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from switch",banco,AdOpenKeySet,AdLockOptimistic
		
If Trim(Request.querystring("switch_data")) = "" Or Trim(Request.querystring("switch_fcg")) = "" Or Trim(Request.querystring("switch_secao")) = "" Or Trim(Request.querystring("switch_esquadrao")) = "" Or Trim(Request.querystring("switch_marca")) = "" Or Trim(Request.querystring("switch_porta")) = "0" Or Trim(Request.querystring("switch_situacao")) = "0" Or Trim(Request.querystring("switch_observa")) = "" Then

		Response.Write("<br><br><p align='center'><font color='#FF0000' size='3'>Você esqueceu de preencher um ou mais campos do formulário.</p>")
		Response.Write("<br><br><p align='center'><font color='#ffffff' size='3'>Use o botão de retornar do navegador para corrigir o erro ou <a href='javascript:history.go(-1)'><font color='#FFFF00' size='3'>clique aqui</a>!</p>")
						
Else

	switch_codigo=request.querystring("switch_codigo")
	switch_data=request.querystring("switch_data")
	switch_fcg=request.querystring("switch_fcg")
	switch_secao=request.querystring("switch_secao")
	switch_esquadrao=request.querystring("switch_esquadrao")	
	switch_marca=request.querystring("switch_marca")
	switch_porta=request.querystring("switch_porta")	
	switch_situacao=request.querystring("switch_situacao")
	switch_observa=request.querystring("switch_observa")
	
	altera = "Update switch set switch_data='"&switch_data&"',switch_fcg='"&switch_fcg&"',switch_secao='"&switch_secao&"',switch_esquadrao='"&switch_esquadrao&"',switch_marca='"&switch_marca&"',switch_porta='"&switch_porta&"',switch_situacao='"&switch_situacao&"',switch_observa='"&switch_observa&"' where switch_codigo="&switch_codigo&" "
	alterar = banco.execute(altera)
	
	rsbanco.movefirst
		While rsbanco("switch_codigo") <> int(switch_codigo)
			rsbanco.movenext
		Wend %>
<body><center><form action="consultas4.asp">
<table border="1" width="600" height="102">
    <tr>
      <td class="fundo1" width="86" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="514" height="102" align="center">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Inventário de Informática</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Cadastro de Switch</b></font>
	  <p style="margin-top: 0; margin-bottom: 0"><font color="#008000"><b>Informações Alteradas com Sucesso!!!</b></font></td>
    </tr>
</table>
<table border="1" width="599" height="1">
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Data</b></td>
        <td class="fundo3" width="210" height="23"><% = rsbanco("switch_data") %></td>
        <td class="fundo1" width="90" height="23" align="center"><b>FCG</b></td>
        <td class="fundo3" width="210" height="23"><% = rsbanco("switch_fcg") %></td>
      </tr>     
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Seção</b></td>
        <td class="fundo3" width="210" height="23"><% = rsbanco("switch_secao") %></td>
		<td class="fundo1" width="90" height="23" align="center"><b>Esquadrão</b></td>
        <td class="fundo3" width="210" height="23"><% = rsbanco("switch_esquadrao") %></td>
      </tr>      
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Marca</b></td>
        <td class="fundo3" width="210" valign="middle" height="23"><% = rsbanco("switch_marca") %></td>      
        <td class="fundo1" width="90" height="23" align="center"><b>Situação</b></td>
        <td class="fundo3" width="210" valign="middle" height="23"><% = rsbanco("switch_situacao") %></td>
      </tr>
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Portas</b></td>
        <td class="fundo3" width="510" height="23" colspan="3"><% = rsbanco("switch_porta") %><b><font size="2" color="#070E5A">&nbsp;&nbsp;&nbsp;Quantidade de portas no Hub ou na Switch.</font></b></td>
      </tr>      
      <tr>
        <td class="fundo1" width="90" height="65" align="center"><b>Observação</b></td>
        <td class="fundo3" width="510" height="65" colspan="3"><% = rsbanco("switch_observa") %></td>
      </tr>                
</table>
<input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;">
<!--#include file="rodape.asp"--></form>
</center></body><% End If %></html>