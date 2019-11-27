<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="styles.asp"-->
<!--#include file="config.asp"-->
<% 	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from estabilizador",banco,AdOpenKeySet,AdLockOptimistic
		
If Trim(Request.querystring("estabilizador_data")) = "" Or Trim(Request.querystring("estabilizador_fcg")) = "" Or Trim(Request.querystring("estabilizador_secao")) = "" Or Trim(Request.querystring("estabilizador_esquadrao")) = "" Or Trim(Request.querystring("estabilizador_marca")) = "0" Or Trim(Request.querystring("estabilizador_situacao")) = "0" Or Trim(Request.querystring("estabilizador_observa")) = "" Then

		Response.Write("<br><br><p align='center'><font color='#FF0000' size='3'>Você esqueceu de preencher um ou mais campos do formulário.</p>")
		Response.Write("<br><br><p align='center'><font color='#ffffff' size='3'>Use o botão de retornar do navegador para corrigir o erro ou <a href='javascript:history.go(-1)'><font color='#FFFF00' size='3'>clique aqui</a>!</p>")
						
Else

	estabilizador_data=request.querystring("estabilizador_data")
	estabilizador_fcg=request.querystring("estabilizador_fcg")
	estabilizador_secao=request.querystring("estabilizador_secao")
	estabilizador_esquadrao=request.querystring("estabilizador_esquadrao")	
	estabilizador_marca=request.querystring("estabilizador_marca")
	estabilizador_situacao=request.querystring("estabilizador_situacao")
	estabilizador_observa=request.querystring("estabilizador_observa")	
	
	rsbanco.AddNew
	rsbanco("estabilizador_data")=estabilizador_data
	rsbanco("estabilizador_fcg")=estabilizador_fcg
	rsbanco("estabilizador_secao")=estabilizador_secao
	rsbanco("estabilizador_esquadrao")=estabilizador_esquadrao	
	rsbanco("estabilizador_marca")=estabilizador_marca
	rsbanco("estabilizador_situacao")=estabilizador_situacao
	rsbanco("estabilizador_observa")=estabilizador_observa	
	rsbanco.Update
	
	rsbanco.movelast %>
<body><center><form action="adicionar6.asp">
<table border="1" width="600" height="102">
    <tr>
      <td class="fundo1" width="86" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="514" height="102" align="center">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Inventário de Informática</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Cadastro de Estabilizador</b></font>
	  <p style="margin-top: 0; margin-bottom: 0"><font color="#008000"><b>Informações Adicionadas com Sucesso!!!</b></font></td>
    </tr>
</table>
<table border="1" width="599" height="1">
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Data</b></td>
        <td class="fundo3" width="210" height="23"><% = rsbanco("estabilizador_data") %></td>
        <td class="fundo1" width="90" height="23" align="center"><b>FCG</b></td>
        <td class="fundo3" width="210" height="23"><% = rsbanco("estabilizador_fcg") %></td>
      </tr>     
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Seção</b></td>
        <td class="fundo3" width="210" height="23"><% = rsbanco("estabilizador_secao") %></td>
		<td class="fundo1" width="90" height="23" align="center"><b>Esquadrão</b></td>
        <td class="fundo3" width="210" height="23"><% = rsbanco("estabilizador_esquadrao") %></td>
      </tr>      
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Marca</b></td>
        <td class="fundo3" width="210" height="23"><% = rsbanco("estabilizador_marca") %></td>      
		<td class="fundo1" width="90" height="23" align="center"><b>Situação</b></td>
        <td class="fundo3" width="210" height="23"><% = rsbanco("estabilizador_situacao") %></td>
      </tr>
      <tr>
        <td class="fundo1" width="90" height="65" align="center"><b>Observação</b></td>
        <td class="fundo3" width="510" height="65" colspan="3"><% = rsbanco("estabilizador_observa") %></td>
      </tr>                
</table>
<input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;">
<!--#include file="rodape.asp"--></form></center></body></html>
<% End If	
	rsbanco.Close		
	banco.Close
	Set rsbanco = Nothing
	Set banco = Nothing %>