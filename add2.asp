<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="styles.asp"-->
<!--#include file="config.asp"-->
<%	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from os",banco,AdOpenKeySet,AdLockOptimistic

If Trim(Request.querystring("os_numero")) = "" Or Trim(Request.querystring("os_numero")) = "" Or Trim(Request.querystring("os_descricaoaber")) = "" Or Trim(Request.querystring("os_militaraber")) = "" Or Trim(Request.querystring("os_ramalaber")) = "" Then

		Response.Write("<br><br><p align='center'><font color='#FF0000' size='3'>Você esqueceu de preencher um ou mais campos do formulário.</p>")
		Response.Write("<br><br><p align='center'><font color='#ffffff' size='3'>Use o botão de retornar do navegador para corrigir o erro ou <a href='javascript:history.go(-1)'><font color='#FFFF00' size='3'>clique aqui</a>!</p>")
						
Else

	set rsbanco1=server.createobject("ADODB.Recordset")
		rsbanco1.open "Select * from os order by os_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic
	
	rsbanco1.MoveLast
		ultimo = rsbanco1("os_codigo")
	rsbanco1.MoveFirst
		numeroos = 1
	
	Do While rsbanco1("os_codigo") <> ultimo
		If rsbanco1("os_numero") > 0 then
			numeroos = numeroos + 1
			rsbanco1.MoveNext
		Else
			rsbanco1.MoveNext
		End If
	Loop 
	
	os_codigo=request.querystring("os_codigo")
	os_numero=numeroos
	os_dataaber=request.querystring("os_dataaber")
	os_descricaoaber=request.querystring("os_descricaoaber")
	os_militaraber=request.querystring("os_militaraber")
	os_ramalaber=request.querystring("os_ramalaber")
	os_status=request.querystring("os_status")
	
	os_status = os_status + 1

	altera = "Update os set os_numero='"&os_numero&"',os_dataaber='"&os_dataaber&"',os_descricaoaber='"&os_descricaoaber&"',os_militaraber='"&os_militaraber&"',os_ramalaber='"&os_ramalaber&"',os_status='"&os_status&"' where os_codigo="&os_codigo&" "
	alterar = banco.execute(altera)
	
	rsbanco.movefirst
		While rsbanco("os_codigo") <> int(os_codigo)
			rsbanco.movenext
		Wend %>
<body><center><form action="consultas3.asp">
<table border="1" width="600" height="102">
    <tr>
      <td class="fundo1" width="90" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="510" height="102" align="center">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Controle de Ordem de Serviço</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Solicitação de abertura de Ordem de Serviço</b></font>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#008000"><b>Ordem de Serviço Aberta com Sucesso!!!</b></font></td>
    </tr>
</table>		
<table border="1" width="600" height="23">
	<tr>
        <td class="fundo1" width="100" height="23" align="center"><b>OS Número</b></td>
        <td class="fundo3" width="200" height="23"><%=rsbanco("os_numero")%></td>
        <td class="fundo1" width="100" height="23" align="center"><b>OS Data</b></td>
        <td class="fundo3" width="200" height="23"><%=rsbanco("os_dataaber")%></font></td>
      </tr>   
      <tr>
        <td class="fundo1" width="100" height="150" align="center"><p align="center"><b>Observações</b></td>
        <td class="fundo3" width="500" height="150" colspan="3"><%=rsbanco("os_descricaoaber")%></font></td>
      </tr>
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>Militar STI</b></td>
        <td class="fundo3" width="200" height="23"><%=rsbanco("os_militaraber")%></font></td>
        <td class="fundo1" width="100" height="23" align="center"><b>Ramal STI</b></td>
        <td class="fundo3" width="200" height="23"><%=rsbanco("os_ramalaber")%></font></td>
	</tr>       
</table>
<input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;">
<!--#include file="rodape.asp"--></form>
</center></body><% End If %></html>