<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="styles.asp"-->
<!--#include file="config.asp"-->
<%	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from os",banco,AdOpenKeySet,AdLockOptimistic

If Trim(Request.querystring("os_dataconc")) = "" Or Trim(Request.querystring("os_observconc")) = "" Or Trim(Request.querystring("os_militarconc")) = "" Or Trim(Request.querystring("os_milrecconc")) = "" Then

		Response.Write("<br><br><p align='center'><font color='#FF0000' size='3'>Você esqueceu de preencher um ou mais campos do formulário.</p>")
		Response.Write("<br><br><p align='center'><font color='#ffffff' size='3'>Use o botão de retornar do navegador para corrigir o erro ou <a href='javascript:history.go(-1)'><font color='#FFFF00' size='3'>clique aqui</a>!</p>")
						
Else

	os_codigo=request.querystring("os_codigo")
	os_dataconc=request.querystring("os_dataconc")
	os_observconc=request.querystring("os_observconc")
	os_militarconc=request.querystring("os_militarconc")
	os_milrecconc=request.querystring("os_milrecconc")
	os_status=request.querystring("os_status")
	
	os_status = os_status + 1

	altera = "Update os set os_dataconc='"&os_dataconc&"',os_observconc='"&os_observconc&"',os_militarconc='"&os_militarconc&"',os_milrecconc='"&os_milrecconc&"',os_status='"&os_status&"' where os_codigo="&os_codigo&" "
	alterar = banco.execute(altera)
	
	rsbanco.movefirst
		While rsbanco("os_codigo") <> int(os_codigo)
			rsbanco.movenext
		Wend %>
<body><center><form action="consultasos2.asp">
<table border="1" width="700" height="102">
    <tr>
      <td class="fundo1" width="90" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="510" height="102" align="center">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Controle de Ordem de Serviço</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Solicitação de abertura de Ordem de Serviço</b></font>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#008000"><b>Ordem de Serviço fechada com sucesso!!!</b></font></td>
    </tr>
</table>		
<table border="1" width="700" height="50">
	<tr>
        <td class="fundo1" width="140" height="50" align="center"><b>Data Fechamento</b></td>
        <td class="fundo3" width="90" height="50"><%=rsbanco("os_dataconc")%></td>
        <td class="fundo1" width="140" height="50" align="center"><b>Militar Fechou</b></td>
        <td class="fundo3" width="95" height="50"><%=rsbanco("os_militarconc")%></td>
        <td class="fundo1" width="140" height="50" align="center"><b>Militar Recebeu</b></td>
        <td class="fundo3" width="95" height="50"><%=rsbanco("os_milrecconc")%></td>
	</tr>
	<tr>
        <td class="fundo1" width="140" height="120" align="center"><b>Observação</b><p><b>de</b></p><p><b>Fechamento</b></td>
        <td class="fundo3" width="560" height="120" colspan="5"><%=rsbanco("os_observconc")%></td>
	</tr> 
</table>
<input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;">
<!--#include file="rodape.asp"--></form>
</center></body><% End If %></html>