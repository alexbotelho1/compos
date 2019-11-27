<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="styles.asp"-->
<!--#include file="config.asp"-->
<% 	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from computador",banco,AdOpenKeySet,AdLockOptimistic
		
If Trim(Request.querystring("computador_data")) = "" Or Trim(Request.querystring("computador_fcg")) = "" Or Trim(Request.querystring("computador_periferico")) = "0" Or Trim(Request.querystring("computador_secao")) = "0" Or Trim(Request.querystring("computador_esquadrao")) = "" Or Trim(Request.querystring("computador_tipo")) = "0" Or Trim(Request.querystring("computador_so")) = "0" Or Trim(Request.querystring("computador_qp")) = "0" Or Trim(Request.querystring("computador_procvelo")) = "" Or Trim(Request.querystring("computador_procfreq")) = "0" Or Trim(Request.querystring("computador_memovelo")) = "" Or Trim(Request.querystring("computador_memocapa")) = "0" Or Trim(Request.querystring("computador_hdtama")) = "" Or Trim(Request.querystring("computador_hdcapa")) = "0" Or Trim(Request.querystring("computador_situacao")) = "0" Or Trim(Request.querystring("computador_observa")) = "" Then

		Response.Write("<br><br><p align='center'><font color='#FF0000' size='3'>Você esqueceu de preencher um ou mais campos do formulário.</p>")
		Response.Write("<br><br><p align='center'><font color='#ffffff' size='3'>Use o botão de retornar do navegador para corrigir o erro ou <a href='javascript:history.go(-1)'><font color='#FFFF00' size='3'>clique aqui</a>!</p>")
						
Else

	computador_data=request.querystring("computador_data")
	computador_periferico=request.querystring("computador_periferico")
	computador_secao=request.querystring("computador_secao")
	computador_esquadrao=request.querystring("computador_esquadrao")	
	computador_fcg=request.querystring("computador_fcg")
	computador_tipo=request.querystring("computador_tipo")
	computador_so=request.querystring("computador_so")
	computador_qp=request.querystring("computador_qp")
	computador_procvelo=request.querystring("computador_procvelo")
	computador_procfreq=request.querystring("computador_procfreq")
	computador_memovelo=request.querystring("computador_memovelo")
	computador_memocapa=request.querystring("computador_memocapa")
	computador_hdtama=request.querystring("computador_hdtama")		
	computador_hdcapa=request.querystring("computador_hdcapa")
	computador_situacao=request.querystring("computador_situacao")
	computador_observa=request.querystring("computador_observa")	
	
	rsbanco.AddNew
	rsbanco("computador_data")=computador_data
	rsbanco("computador_periferico")=computador_periferico
	rsbanco("computador_secao")=computador_secao
	rsbanco("computador_esquadrao")=computador_esquadrao	
	rsbanco("computador_fcg")=computador_fcg
	rsbanco("computador_tipo")=computador_tipo
	rsbanco("computador_so")=computador_so
	rsbanco("computador_qp")=computador_qp	
	rsbanco("computador_procvelo")=computador_procvelo
	rsbanco("computador_procfreq")=computador_procfreq
	rsbanco("computador_memovelo")=computador_memovelo
	rsbanco("computador_memocapa")=computador_memocapa
	rsbanco("computador_hdtama")=computador_hdtama
	rsbanco("computador_hdcapa")=computador_hdcapa
	rsbanco("computador_situacao")=computador_situacao
	rsbanco("computador_observa")=computador_observa	
	rsbanco.Update
	
	rsbanco.movelast %>
<body><center><form action="adicionar4.asp">
<table border="1" width="600" height="102">
    <tr>
      <td class="fundo1" width="90" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="510" height="102" align="center">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Inventário de Informática</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Cadastro de Hardware</b></font>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#008000"><b>Informações Inseridas com Sucesso!!!</b></font></td>
    </tr>
</table>  
<table border="1" width="600" height="1">
      <tr>
        <td class="fundo1" width="96" height="23" align="center"><b>Data</b></td>
        <td class="fundo3" width="147" height="23"><%=rsbanco("computador_data")%></td>
        <td class="fundo1" width="116" height="23" align="center"><b>Periférico</b></td>
        <td class="fundo3" width="372" height="23"><%=rsbanco("computador_periferico")%></td>
      </tr>     
      <tr>
        <td class="fundo1" width="96" height="23" align="center"><b>Seção</b></td>
        <td class="fundo3" width="317" height="23"><%=rsbanco("computador_secao")%></td>
		<td class="fundo1" width="116" height="23" align="center"><b>Esquadrão</b></td>
        <td class="fundo3" width="372" height="23"><%=rsbanco("computador_esquadrao")%></td>
      </tr>      
      <tr>
        <td class="fundo1" width="96" height="23" align="center"><b>FCG</b></td>
        <td class="fundo3" width="317" height="23"><%=rsbanco("computador_fcg")%></td>
        <td class="fundo1" width="116" height="23" align="center"><b>Tipo</b></td>
        <td class="fundo3" width="372" height="23"><%=rsbanco("computador_tipo")%></tr>
      <tr>
        <td class="fundo1" width="96" height="23" align="center"><b>Sist Oper</b></td>
        <td class="fundo3" width="317" height="23"><%=rsbanco("computador_so")%></td>
        <td class="fundo1" width="116" height="23" align="center"><b>Qtd Proces</b></td>
        <td class="fundo3" width="372" height="23"><%=rsbanco("computador_qp")%></td>
      </tr>
      <tr>
        <td class="fundo1" width="96" height="23" align="center"><b>Processador</b></td>
        <td class="fundo3" width="317" height="23"><%=rsbanco("computador_procvelo")%>  <%=rsbanco("computador_procfreq")%></td>
        <td class="fundo1" width="116" height="23" align="center"><b>Memória</b></td>
        <td class="fundo3" width="372" height="23"><%=rsbanco("computador_memovelo")%>  <%=rsbanco("computador_memocapa")%></td>
      </tr>        
      <tr>
        <td class="fundo1" width="96" height="23" align="center"><b>Hard Disk</b></td>
        <td class="fundo3" width="317" height="23"><%=rsbanco("computador_hdtama")%>  <%=rsbanco("computador_hdcapa")%></td>
        <td class="fundo1" width="116" height="23" align="center"><b>Situação</b></td>
        <td class="fundo3" width="372" height="23"><%=rsbanco("computador_situacao")%></td>
      </tr>
      <tr>
        <td class="fundo1" width="91" height="65" align="center"><b>Observação</b></td>
        <td class="fundo3" width="510" height="65" colspan="3"><%=rsbanco("computador_observa")%></td>
     </tr>                
</table>
<input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;">
<!--#include file="rodape.asp"--></form></center></body></html>
<% End If	
	rsbanco.Close		
	banco.Close
	Set rsbanco = Nothing
	Set banco = Nothing %>