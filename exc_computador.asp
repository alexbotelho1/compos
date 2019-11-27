<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% If Session("LoggedIn") = True Then	
	codigo = request.querystring("codigo")

	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from computador WHERE computador_codigo = "&codigo&"",banco,AdOpenKeySet,AdLockOptimistic
%>
<body><center><form action="exc3.asp">
<table border="1" width="600" height="102">
    <tr>
      <td class="fundo1" width="90" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="510" height="102" align="center">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Inventário de Informática</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Excluir Computador</b></font></td>
    </tr>
</table>
<input type=hidden name=computador_codigo value="<%=rsbanco("computador_codigo")%>">
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
        <td class="fundo3" width="317" height="23"><%=rsbanco("computador_procvelo")%>&nbps;<%=rsbanco("computador_procfreq")%></td>
        <td class="fundo1" width="116" height="23" align="center"><b>Memória</b></td>
        <td class="fundo3" width="372" height="23"><%=rsbanco("computador_memovelo")%>&nbsp;<%=rsbanco("computador_memocapa")%></td>
      </tr>        
      <tr>
        <td class="fundo1" width="96" height="23" align="center"><b>Hard Disk</b></td>
        <td class="fundo3" width="317" height="23"><%=rsbanco("computador_hdtama")%>&nbsp;<%=rsbanco("computador_hdcapa")%></td>
        <td class="fundo1" width="116" height="23" align="center"><b>Situação</b></td>
        <td class="fundo3" width="372" height="23"><%=rsbanco("computador_situacao")%></td>
      </tr>
      <tr>
        <td class="fundo1" width="91" height="65" align="center"><b>Observação</b></td>
        <td class="fundo3" width="510" height="65" colspan="3"><%=rsbanco("computador_observa")%></td>
     </tr>                
</table>
<input type="submit" value="&nbsp;&nbsp;Excluir&nbsp;&nbsp;" name="BTincluir"></form>
<form action="consultas4.asp"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;">
<!--#include file="rodape.asp"--></form></body></html>
<% Else
	Response.Redirect("admin.asp")
End If %>