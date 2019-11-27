<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="styles.asp"-->
<!--#include file="config.asp"-->
<%	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from os",banco,AdOpenKeySet,AdLockOptimistic

If Trim(Request.querystring("os_tempoexec")) = "" Or Trim(Request.querystring("os_dataexec")) = "" Or Trim(Request.querystring("os_descricaoexec")) = "" Or Trim(Request.querystring("os_matusadoexec")) = "" Or Trim(Request.querystring("os_militarexec")) = "" Then

		Response.Write("<br><br><p align='center'><font color='#FF0000' size='3'>Você esqueceu de preencher um ou mais campos do formulário.</p>")
		Response.Write("<br><br><p align='center'><font color='#ffffff' size='3'>Use o botão de retornar do navegador para corrigir o erro ou <a href='javascript:history.go(-1)'><font color='#FFFF00' size='3'>clique aqui</a>!</p>")
						
Else
	os_codigo=request.querystring("os_codigo")
	os_tempoexec=request.querystring("os_tempoexec")
	os_dataexec=request.querystring("os_dataexec")
	os_descricaoexec=request.querystring("os_descricaoexec")
	os_matusadoexec=request.querystring("os_matusadoexec")
	os_militarexec=request.querystring("os_militarexec")
	os_status=request.querystring("os_status")
	
	os_status = os_status + 1

	altera = "Update os set os_tempoexec='"&os_tempoexec&"',os_dataexec='"&os_dataexec&"',os_descricaoexec='"&os_descricaoexec&"',os_matusadoexec='"&os_matusadoexec&"',os_militarexec='"&os_militarexec&"',os_status='"&os_status&"' where os_codigo="&os_codigo&" "
	alterar = banco.execute(altera)
	
	rsbanco.movefirst
		While rsbanco("os_codigo") <> int(os_codigo)
			rsbanco.movenext
		Wend %>
<body><center><form action="consultasos2.asp">
<table border="1" width="600" height="102">
    <tr>
      <td class="fundo1" width="90" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="510" height="102" align="center">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Controle de Ordem de Serviço</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Solicitação de abertura de Ordem de Serviço</b></font>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#008000"><b>Ordem de Serviço Executada, incluída com sucesso!!!</b></font></td>
    </tr>
</table>		
<table border="1" width="600" height="23">
	<tr>
        <td class="fundo1" width="100" height="23" align="center"><b>Tempo H/H</b></td>
        <td class="fundo3" width="250" height="23"><%=rsbanco("os_tempoexec")%></td>
        <td class="fundo1" width="100" height="23" align="center"><b>Data Exec.</b></td>
        <td class="fundo3" width="250" height="23"><%=rsbanco("os_dataexec")%></td>
      </tr>
      <tr>
        <td class="fundo1" width="100" height="120" align="center"><b>Observações</b></td>
        <td class="fundo3" width="500" height="120" colspan="3"><%=rsbanco("os_descricaoexec")%></td>
      </tr>       
      <tr>
        <td class="fundo1" width="100" height="120" align="center"><b>Observações</b></td>
        <td class="fundo3" width="500" height="120" colspan="3"><%=rsbanco("os_matusadoexec")%></td>
      </tr>
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>Militar Exec.</b></td>
        <td class="fundo3" width="500" height="23" colspan="3"><%=rsbanco("os_militarexec")%></td>
	</tr>     
</table>
<input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;">
<!--#include file="rodape.asp"--></form>
</center></body><% End If %></html>