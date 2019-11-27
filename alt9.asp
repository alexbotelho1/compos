<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="styles.asp"-->
<!--#include file="config.asp"-->
<%	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from os",banco,AdOpenKeySet,AdLockOptimistic

If Trim(Request.querystring("os_solicdata")) = "" Or Trim(Request.querystring("os_solicperiferico")) = "" Or Trim(Request.querystring("os_solicdescricao")) = "" Or Trim(Request.querystring("os_solicmilitar")) = "" Or Trim(Request.querystring("os_solicesquadrao")) = "" Or Trim(Request.querystring("os_solicsecao")) = "" Or Trim(Request.querystring("os_solicramal")) = "" Or Trim(Request.querystring("os_numero")) = "" Or Trim(Request.querystring("os_dataaber")) = "" Or Trim(Request.querystring("os_descricaoaber")) = "" Or Trim(Request.querystring("os_militaraber")) = "" Or Trim(Request.querystring("os_ramalaber")) = "" Or Trim(Request.querystring("os_tempoexec")) = "" Or Trim(Request.querystring("os_dataexec")) = "" Or Trim(Request.querystring("os_descricaoexec")) = "" Or Trim(Request.querystring("os_matusadoexec")) = "" Or Trim(Request.querystring("os_militarexec")) = "" Then

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
	os_numero=request.querystring("os_numero")
	os_dataaber=request.querystring("os_dataaber")
	os_descricaoaber=request.querystring("os_descricaoaber")
	os_militaraber=request.querystring("os_militaraber")
	os_ramalaber=request.querystring("os_ramalaber")
	os_tempoexec=request.querystring("os_tempoexec")
	os_dataexec=request.querystring("os_dataexec")
	os_descricaoexec=request.querystring("os_descricaoexec")
	os_matusadoexec=request.querystring("os_matusadoexec")
	os_militarexec=request.querystring("os_militarexec")
	
	altera = "Update os set os_solicdata='"&os_solicdata&"',os_solicperiferico='"&os_solicperiferico&"',os_solicdescricao='"&os_solicdescricao&"',os_solicmilitar='"&os_solicmilitar&"',os_solicesquadrao='"&os_solicesquadrao&"',os_solicsecao='"&os_solicsecao&"',os_solicramal='"&os_solicramal&"',os_numero='"&os_numero&"',os_dataaber='"&os_dataaber&"',os_descricaoaber='"&os_descricaoaber&"',os_militaraber='"&os_militaraber&"',os_ramalaber='"&os_ramalaber&"',os_tempoexec='"&os_tempoexec&"',os_dataexec='"&os_dataexec&"',os_descricaoexec='"&os_descricaoexec&"',os_matusadoexec='"&os_matusadoexec&"',os_militarexec='"&os_militarexec&"' where os_codigo="&os_codigo&" "
	alterar = banco.execute(altera)
	
	rsbanco.movefirst
		While rsbanco("os_codigo") <> int(os_codigo)
			rsbanco.movenext
		Wend %>
<body><center><form method="GET" action="consultasos2.asp">
<table border="1" width="600" height="102">
    <tr>
      <td class="fundo1" width="100" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="500" height="102" align="center">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Controle de Ordem de Serviço</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#008000"><b>Ordem de Serviço editada com sucesso!!!</b></font></td>
    </tr>
</table>
<table border="1" width="600" height="23">
      <tr>
        <td class="fundo2" width="600" height="23" align="center" colspan="4"><font color="#008000"><b>Formulário de Solicitação</b></font></td>
      </tr>
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>Número</b></td>
        <td class="fundo3" width="500" height="23" colspan="3"><%=rsbanco("os_codigo")%></td>
      </tr>   
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>Data Solic</b></td>
        <td class="fundo3" width="150" height="23"><%=rsbanco("os_solicdata")%></td>
        <td class="fundo1" width="100" height="23" align="center"><b>Periférico</b></td>
        <td class="fundo3" width="150" height="23"><%=rsbanco("os_solicperiferico")%></td>
      </tr>      
      <tr>
        <td class="fundo1" width="100" height="197" align="center"><b>Descrição</b></p><p align="center"><b>do</b></p><p align="center"><b>Problema</b></td>
        <td class="fundo3" width="500" height="197" colspan="3"><%=rsbanco("os_solicdescricao")%></td>
      </tr>
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>Solicitante</b></td>
        <td class="fundo3" width="150" height="23"><%=rsbanco("os_solicmilitar")%></td>
        <td class="fundo1" width="100" height="23" align="center"><b>Esquadrão</b></td>
        <td class="fundo3" width="150" height="23"><%=rsbanco("os_solicesquadrao")%></td>
      </tr>
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>Seção</b></td>
        <td class="fundo3" width="150" height="23"><%=rsbanco("os_solicsecao")%></td>
        <td class="fundo1" width="100" height="23" align="center"><b>Ramal</b></td>
        <td class="fundo3" width="150" height="23"><%=rsbanco("os_solicramal")%></td>
      </tr>       
    </table>
<table border="1" width="600" height="23">
      <tr>
        <td class="fundo2" width="600" height="23" align="center" colspan="4"><font color="#008000"><b>Formulário de Abertura de Ordem de Serviço</b></font></td>
      </tr>    
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>OS Número</b></td>      
        <td class="fundo3" width="150" height="23"><%=rsbanco("os_numero")%></td>
        <td class="fundo1" width="100" height="23" align="center"><b>OS Data</b></td>
        <td class="fundo3" width="150" height="23"><%=rsbanco("os_dataaber")%></td>
      </tr>    
      <tr>
        <td class="fundo1" width="100" height="197" align="center"><p><b>Observações</b></td>
        <td class="fundo3" width="500" height="197" colspan="3"><%=rsbanco("os_descricaoaber")%></td>
      </tr>
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>Militar STI</b></td>
        <td class="fundo3" width="150" height="23"><%=rsbanco("os_militaraber")%></td>
        <td class="fundo1" width="100" height="23" align="center"><b>Ramal STI</b></td>
        <td class="fundo3" width="150" height="23"><%=rsbanco("os_ramalaber")%></td>
      </tr>       
    </table>  
    <table border="1" width="600" height="23">
      <tr>
        <td class="fundo2" width="600" height="23" align="center" colspan="4"><font color="#008000"><b>Formulário de Execução de Ordem de Serviço</b></font></td>
      </tr>      
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>Tempo H/H</b></td>
        <td class="fundo3" width="150" height="23"><%=rsbanco("os_tempoexec")%></td>
        <td class="fundo1" width="100" height="23" align="center"><b>Data Exec.</b></td>
        <td class="fundo3" width="150" height="23"><%=rsbanco("os_dataexec")%></td>
      </tr>    
      <tr>
        <td class="fundo1" width="100" height="134" align="center"><b>Serviço</b><p align="center"><b>Executado</b></td>
        <td class="fundo3" width="500" height="134" colspan="3"><%=rsbanco("os_descricaoexec")%></td>
      </tr>
      <tr>
        <td class="fundo1" width="100" height="121" align="center"><b>Material</b><p align="center"><b>Utilizado</b></td>
        <td class="fundo3" width="500" height="121" colspan="3"><%=rsbanco("os_matusadoexec")%></td>
      </tr>      
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>Militar Exec.</b></td>
        <td class="fundo3" width="500" height="23" colspan="3"><%=rsbanco("os_militarexec")%></td>
      </tr>       
    </table>
<input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;">
<!--#include file="rodape.asp"--></form>
</center></body><% End If %></html>