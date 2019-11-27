<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!-- #include file="config.asp" -->
<!--#include file="styles.asp"-->
<%	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from os",banco,AdOpenKeySet,AdLockOptimistic
		
If Trim(Request.querystring("os_solicdata")) = "" Or Trim(Request.querystring("os_solicperiferico")) = "" Or Trim(Request.querystring("os_solicdescricao")) = "" Or Trim(Request.querystring("os_solicmilitar")) = "" Or Trim(Request.querystring("os_solicesquadrao")) = "" Or Trim(Request.querystring("os_solicsecao")) = "" Or Trim(Request.querystring("os_solicramal")) = "" Then

		Response.Write("<br><br><p align='center'><font color='#FF0000' size='3'>Você esqueceu de preencher um ou mais campos do formulário.</p>")
		Response.Write("<br><br><p align='center'><font color='#ffffff' size='3'>Use o botão de retornar do navegador para corrigir o erro ou <a href='javascript:history.go(-1)'><font color='#FFFF00' size='3'>clique aqui</a>!</p>")
						
Else

	os_codigo=request.querystring("os_codigo")
	os_solicdata=request.querystring("os_solicdata")
	os_solicperiferico=request.querystring("os_solicperiferico")
	os_solicdescricao=request.querystring("os_solicdescricao")
	os_solicmilitar=request.querystring("os_solicmilitar")
	os_solicesquadrao=request.querystring("os_solicesquadrao")
	os_solicsecao=request.querystring("os_solicsecao")
	os_solicramal=request.querystring("os_solicramal")
	
	altera = "Update os set os_solicdata='"&os_solicdata&"',os_solicperiferico='"&os_solicperiferico&"',os_solicdescricao='"&os_solicdescricao&"',os_solicmilitar='"&os_solicmilitar&"',os_solicesquadrao='"&os_solicesquadrao&"',os_solicsecao='"&os_solicsecao&"',os_solicramal='"&os_solicramal&"' where os_codigo="&os_codigo&" "
	alterar = banco.execute(altera)
	
	rsbanco.movefirst
		While rsbanco("os_codigo") <> int(os_codigo)
			rsbanco.movenext
		Wend %>
<body><center><form action="consultas_completa2.asp">
<table border="1" width="600" height="102">
    <tr>
      <td class="fundo1" width="100" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="500" height="102" align="center">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Controle de Ordem de Serviço</b></font>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Consulta das solicitações de abertura de Ordem de Serviço</b></font>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#008000"><b>Solicitação Editada com Sucesso!!!</b></font></td>
    </tr>
</table>
<table border="1" width="600" height="23">
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>Número</b></td>
        <td class="fundo3" width="500" height="23" colspan="3"><%=rsbanco("os_codigo")%></td>
      </tr>
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>Data</b></td>
        <td class="fundo3" width="200" height="23"><%=rsbanco("os_solicdata")%></td>
        <td class="fundo1" width="100" height="23" align="center"><b>Periférico</b></td>
        <td class="fundo3" width="200" height="23"><%=rsbanco("os_solicperiferico")%></td>
      </tr>      
      <tr>
        <td class="fundo1" width="100" height="146" align="center"><b>Descrição</b></p><p align="center"><b>do</b></p><p align="center"><b>Problema</b></td>
        <td class="fundo3" width="500" height="146" colspan="3"><%=rsbanco("os_solicdescricao")%></td>
      </tr>    
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>Solicitante</b></td>
        <td class="fundo3" width="200" height="23"><%=rsbanco("os_solicmilitar")%></td>
        <td class="fundo1" width="100" height="23" align="center"><b>Esquadrão</b></td>
        <td class="fundo3" width="200" height="23"><%=rsbanco("os_solicesquadrao")%></td>
      </tr>
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>Seção</b></td>
        <td class="fundo3" width="200" height="23"><%=rsbanco("os_solicsecao")%></td>
        <td class="fundo1" width="100" height="23" align="center"><b>Ramal</b></td>
        <td class="fundo3" width="200" height="23"><%=rsbanco("os_solicramal")%></td>
      </tr> 
    </table>
<input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;">
<!--#include file="rodape.asp"--></form>
</center></body><% End If %></html>