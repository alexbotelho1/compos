<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!-- #include file="config.asp" -->
<!--#include file="styles.asp"-->
<%	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from impressora",banco,AdOpenKeySet,AdLockOptimistic
		
If Trim(Request.querystring("impressora_data")) = "" Or Trim(Request.querystring("impressora_fcg")) = "" Or Trim(Request.querystring("impressora_secao")) = "" Or Trim(Request.querystring("impressora_esquadrao")) = "" Or Trim(Request.querystring("impressora_marca")) = "" Or Trim(Request.querystring("impressora_modelo")) = "" Or Trim(Request.querystring("impressora_impressao")) = "" Or Trim(Request.querystring("impressora_cor")) = "" Or Trim(Request.querystring("impressora_consumo")) = "" Or Trim(Request.querystring("impressora_situacao")) = "" Or Trim(Request.querystring("impressora_observa")) = "" Then

		Response.Write("<br><br><p align='center'><font color='#FF0000' size='3'>Você esqueceu de preencher um ou mais campos do formulário.</p>")
		Response.Write("<br><br><p align='center'><font color='#ffffff' size='3'>Use o botão de retornar do navegador para corrigir o erro ou <a href='javascript:history.go(-1)'><font color='#FFFF00' size='3'>clique aqui</a>!</p>")
						
Else

	impressora_codigo=request.querystring("impressora_codigo")
	impressora_data=request.querystring("impressora_data")
	impressora_fcg=request.querystring("impressora_fcg")
	impressora_secao=request.querystring("impressora_secao")
	impressora_esquadrao=request.querystring("impressora_esquadrao")	
	impressora_marca=request.querystring("impressora_marca")
	impressora_modelo=request.querystring("impressora_modelo")
	impressora_impressao=request.querystring("impressora_impressao")
	impressora_cor=request.querystring("impressora_cor")
	impressora_colorido=request.querystring("impressora_colorido")
	impressora_preto=request.querystring("impressora_preto")
	impressora_toner=request.querystring("impressora_toner")
	impressora_consumo=request.querystring("impressora_consumo")
	impressora_situacao=request.querystring("impressora_situacao")
	impressora_observa=request.querystring("impressora_observa")
	
	altera = "Update impressora set	impressora_data='"&impressora_data&"',impressora_fcg='"&impressora_fcg&"',impressora_secao='"&impressora_secao&"',impressora_esquadrao='"&impressora_esquadrao&"',impressora_marca='"&impressora_marca&"',impressora_modelo='"&impressora_modelo&"',impressora_impressao='"&impressora_impressao&"',impressora_cor='"&impressora_cor&"',impressora_colorido='"&impressora_colorido&"',impressora_preto='"&impressora_preto&"',impressora_toner='"&impressora_toner&"',impressora_consumo='"&impressora_consumo&"',impressora_situacao='"&impressora_situacao&"',impressora_observa='"&impressora_observa&"' where impressora_codigo="&impressora_codigo&" "
	alterar = banco.execute(altera)
	
	rsbanco.movefirst
		While rsbanco("impressora_codigo") <> int(impressora_codigo)
			rsbanco.movenext
		Wend %>
<body><center><form action="consultas4.asp">
<table border="1" width="600" height="102">
    <tr>
      <td class="fundo1" width="86" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="514" height="102" align="center">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Inventário de Informática</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Cadastro de Impressora</b></font>
	  <p style="margin-top: 0; margin-bottom: 0"><font color="#008000"><b>Informações Alteradas com Sucesso!!!</b></font></td>
    </tr>
</table>
<table border="1" width="599" height="1">
      <tr>
        <td class="fundo1" width="90" height="23" align="center"><b>Data</b></td>
        <td class="fundo3" width="210" height="23"><% = rsbanco("impressora_data") %></td>
        <td class="fundo1" width="90" height="23" align="center"><b>FCG</b></td>
        <td class="fundo3" width="210" height="23"><% = rsbanco("impressora_fcg") %></td>
      </tr>     
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
        <td class="fundo3" width="510" height="65" colspan="3"><% = rsbanco("impressora_observa") %></td>
      </tr>                
</table>
<input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;">
<!--#include file="rodape.asp"--></form>
</center></body><% End If %></html>