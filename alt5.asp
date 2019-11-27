<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!-- #include file="config.asp" -->
<!--#include file="styles.asp"-->
<%	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from nobreak",banco,AdOpenKeySet,AdLockOptimistic
		
If Trim(Request.querystring("nobreak_data")) = "" Or Trim(Request.querystring("nobreak_fcg")) = "" Or Trim(Request.querystring("nobreak_secao")) = "" Or Trim(Request.querystring("nobreak_esquadrao")) = "" Or Trim(Request.querystring("nobreak_marca")) = "" Or Trim(Request.querystring("nobreak_potencia")) = "0" Or Trim(Request.querystring("nobreak_saida")) = "0" Or Trim(Request.querystring("nobreak_situacao")) = "0" Or Trim(Request.querystring("nobreak_observa")) = "" Then

		Response.Write("<br><br><p align='center'><font color='#FF0000' size='3'>Você esqueceu de preencher um ou mais campos do formulário.</p>")
		Response.Write("<br><br><p align='center'><font color='#ffffff' size='3'>Use o botão de retornar do navegador para corrigir o erro ou <a href='javascript:history.go(-1)'><font color='#FFFF00' size='3'>clique aqui</a>!</p>")
						
Else

	nobreak_codigo=request.querystring("nobreak_codigo")
	nobreak_data=request.querystring("nobreak_data")
	nobreak_fcg=request.querystring("nobreak_fcg")
	nobreak_secao=request.querystring("nobreak_secao")
	nobreak_esquadrao=request.querystring("nobreak_esquadrao")	
	nobreak_marca=request.querystring("nobreak_marca")
	nobreak_potencia=request.querystring("nobreak_potencia")
	nobreak_saida=request.querystring("nobreak_saida")
	nobreak_situacao=request.querystring("nobreak_situacao")
	nobreak_observa=request.querystring("nobreak_observa")
	
	altera = "Update nobreak set nobreak_data='"&nobreak_data&"',nobreak_fcg='"&nobreak_fcg&"',nobreak_secao='"&nobreak_secao&"',nobreak_esquadrao='"&nobreak_esquadrao&"',nobreak_marca='"&nobreak_marca&"',nobreak_potencia='"&nobreak_potencia&"',nobreak_saida='"&nobreak_saida&"',nobreak_situacao='"&nobreak_situacao&"',nobreak_observa='"&nobreak_observa&"' where nobreak_codigo="&nobreak_codigo&" "
	alterar = banco.execute(altera)
	
	rsbanco.movefirst
		While rsbanco("nobreak_codigo") <> int(nobreak_codigo)
			rsbanco.movenext
		Wend %>
<body><center><form action="consultas4.asp">
<table border="1" width="600" height="102">
    <tr>
      <td class="fundo1" width="86" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="514" height="102" align="center">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Inventário de Informática</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Cadastro de Nobreak</b></font>
	  <p style="margin-top: 0; margin-bottom: 0"><font color="#008000"><b>Informações Alteradas com Sucesso!!!</b></font></td>
    </tr>
</table>
<table border="1" width="599" height="1">
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Data</b></td>
        <td class="fundo3" width="210" height="23"><% = rsbanco("nobreak_data") %></td>
        <td class="fundo1" width="90" height="23" align="center"><b>FCG</b></td>
        <td class="fundo3" width="210" height="23"><% = rsbanco("nobreak_fcg") %></td>
      </tr>     
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Seção</b></td>
        <td class="fundo3" width="210" height="23"><% = rsbanco("nobreak_secao") %></td>
		<td class="fundo1" width="90" height="23" align="center"><b>Esquadrão</b></td>
        <td class="fundo3" width="210" height="23"><% = rsbanco("nobreak_esquadrao") %></td>
      </tr>      
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Marca</b></td>
        <td class="fundo3" width="210" valign="middle" height="23"><% = rsbanco("nobreak_marca") %></td>      
        <td class="fundo1" width="90" height="23" align="center"><b>Potência</b></td>
        <td class="fundo3" width="210" height="23"><% = rsbanco("nobreak_potencia") %> KVa</td>
      </tr>    
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Qtd Saída</b></td>
        <td class="fundo3" width="210" height="23"><% = rsbanco("nobreak_saida") %></td>
        <td class="fundo1" width="90" height="23" align="center"><b>Situação</b></td>
        <td class="fundo3" width="210" valign="middle" height="23"><% = rsbanco("nobreak_situacao") %></td>
      </tr>
      <tr>
        <td class="fundo1" width="90" height="65" align="center"><b>Observação</b></td>
        <td class="fundo3" width="510" height="65" colspan="3"><% = rsbanco("nobreak_observa") %></td>
      </tr>                
</table>
<input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;">
<!--#include file="rodape.asp"--></form>
</center></body><% End If %></html>