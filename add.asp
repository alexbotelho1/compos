<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="styles.asp"-->
<!--#include file="config.asp"-->
<% 	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from os",banco,AdOpenKeySet,AdLockOptimistic
		
	ip = Request.ServerVariables("REMOTE_ADDR")
	host = Request.ServerVariables("REMOTE_HOST")
	logon = Request.ServerVariables("LOGON_USER")
	serverorigem = Request.ServerVariables("SERVER_NAME")		
		
If Trim(Request.querystring("os_solicdata")) = "" Or Trim(Request.querystring("os_solicperiferico")) = "0" Or Trim(Request.querystring("os_solicdescricao")) = "" Or Trim(Request.querystring("os_solicmilitar")) = "" Or Trim(Request.querystring("os_solicesquadrao")) = "0" Or Trim(Request.querystring("os_solicsecao")) = "0" Or Trim(Request.querystring("os_solicramal")) = "" Then

		Response.Write("<br><br><p align='center'><font color='#FF0000' size='3'>Você esqueceu de preencher um ou mais campos do formulário.</p>")
		Response.Write("<br><br><p align='center'><font color='#ffffff' size='3'>Use o botão de retornar do navegador para corrigir o erro ou <a href='javascript:history.go(-1)'><font color='#FFFF00' size='3'>clique aqui</a>!</p>")
						
Else

	os_dia=request.querystring("os_dia")
	os_mes=request.querystring("os_mes")
	os_ano=request.querystring("os_ano")
	os_solicdata=request.querystring("os_solicdata")
	os_solicperiferico=request.querystring("os_solicperiferico")
	os_solicdescricao=request.querystring("os_solicdescricao")
	os_solicmilitar=request.querystring("os_solicmilitar")
	os_solicesquadrao=request.querystring("os_solicesquadrao")
	os_solicsecao=request.querystring("os_solicsecao")
	os_solicramal=request.querystring("os_solicramal")
	os_ip=ip
	os_host=host
	os_logon=logon
	os_server=serverorigem 

	rsbanco.AddNew
	rsbanco("os_dia")=os_dia
	rsbanco("os_mes")=os_mes
	rsbanco("os_ano")=os_ano			
	rsbanco("os_solicdata")=os_solicdata
	rsbanco("os_solicperiferico")=os_solicperiferico
	rsbanco("os_solicdescricao")=os_solicdescricao
	rsbanco("os_solicmilitar")=os_solicmilitar
	rsbanco("os_solicesquadrao")=os_solicesquadrao
	rsbanco("os_solicsecao")=os_solicsecao
	rsbanco("os_solicramal")=os_solicramal
	rsbanco("os_ip")=os_ip
	rsbanco("os_host")=os_host	
	rsbanco("os_logon")=os_logon
	rsbanco("os_server")=os_server	
	rsbanco.Update
	
	rsbanco.movelast %>
<body><center><form action="adicionar.asp">
<table border="1" width="600" height="102">
    <tr>
      <td class="fundo1" width="90" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="510" height="102" align="center">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Controle de Ordem de Serviço</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Solicitação de abertura de Ordem de Serviço</b></font>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#008000"><b>Solicitação Inserida com Sucesso!!!</b></font></td>
    </tr>
</table>
<table border="1" width="600" height="319">
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>Número</b></td>
        <td class="fundo3" width="500" height="23" align="center" colspan="3">POR FAVOR GUARDE ESSE NÚMERO: <font color="#FF0000"><b><%=rsbanco("os_codigo")%></b></font></td>
      </tr>    
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>Data</b></td>
        <td class="fundo3" width="200" height="23" align="center"><%=rsbanco("os_solicdata")%></font></td>
        <td class="fundo1" width="100" height="23" align="center"><b>Periférico</b></td>
        <td class="fundo3" width="200" height="23" align="center"><%=rsbanco("os_solicperiferico")%></font></td>
      </tr>      
      <tr>
        <td class="fundo1" width="100" height="197" align="center"><b>Descrição</b></p><p align="center"><b>do</b></p><p align="center"><b>Problema</b></td>
        <td class="fundo3" width="500" height="197" colspan="3"><%=rsbanco("os_solicdescricao")%></font></td>
      </tr>
      <tr>
        <td class="fundo1" class="fundo1" width="100" height="23" align="center"><b>Solicitante</b></td>
        <td class="fundo3" width="200" height="23" align="center"><%=rsbanco("os_solicmilitar")%></font></td>
        <td class="fundo1" width="100" height="23" align="center"><b>Esquadrão</b></td>
        <td class="fundo3" width="200" height="23" align="center"><%=rsbanco("os_solicesquadrao")%></font></td>
      </tr>
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>Seção</b></td>
        <td class="fundo3" width="200" height="23" align="center"><%=rsbanco("os_solicsecao")%></font></td>
        <td class="fundo1" width="100" height="23" align="center"><b>Ramal</b></td>
        <td class="fundo3" width="200" height="23" align="center"><%=rsbanco("os_solicramal")%></font></td>
      </tr>       
    </table>
<input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;">
<!--#include file="rodape.asp"--></form>
</center></body></html>
<% End If	
	rsbanco.Close		
	banco.Close
	Set rsbanco = Nothing
	Set banco = Nothing %>