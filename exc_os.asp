<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% If Session("LoggedIn") = True Then	
	codsolic = request.querystring("codsolic")

	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from os WHERE os_codigo = "&codsolic&"",banco,AdOpenKeySet,AdLockOptimistic
%>
<body><center><form method="GET" action="exc2.asp">
<table border="1" width="700" height="102">
    <tr>
      <td class="fundo1" width="100" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="600" height="102" align="center">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Controle de Ordem de Serviço</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Excluir Ordem de Serviço</b></font></td>
    </tr>
</table>
<table border="1" width="700" height="410">
    <tr>
      <td class="fundo1" width="94" height="34" align="center"><b><font size="2">Os Número</font></b></td>
      <td class="fundo3" width="110" height="34" align="center"><%=rsbanco("os_numero")%></td>
      <td class="fundo1" width="102" height="34" align="center"><b><font size="2">Data Solicitação</font></b></td>
      <td class="fundo3" width="130" height="34" align="center"><font size="2"><%=rsbanco("os_solicdata")%></td>
      <td class="fundo1" width="109" height="34" align="center"><b><font size="2">Solicitado Por</font></b></td>
      <td class="fundo3" width="141" height="34" align="center"><%=rsbanco("os_solicmilitar")%></td>
    </tr>
    <tr class="fundo1">
      <td width="206" height="26" colspan="2" align="center"><b><font size="2">Observações (Abertura)</font></b></td>
      <td width="234" height="26" colspan="2" align="center"><b><font size="2">Descrição do Serviço</font></b></td>
      <td width="252" height="26" colspan="2" align="center"><b><font size="2">Observações (Conclusão)</font></b></td>
    </tr>
    <tr class="fundo3">
      <td width="206" height="121" colspan="2"><%=rsbanco("os_descricaoaber")%></td>
      <td width="234" height="121" colspan="2"><%=rsbanco("os_descricaoexec")%></td>
      <td width="252" height="121" colspan="2"><%=rsbanco("os_observconc")%></td>
    </tr>
    <tr>
      <td class="fundo1" width="94" height="41" align="center"><b><font size="2">Militar STI</font></b></td>
      <td class="fundo3" width="110" height="41" align="center"><%=rsbanco("os_militaraber")%></td>
      <td class="fundo1" width="102" height="41" align="center"><b><font size="2">Militar Exec.</font></b></td>
      <td class="fundo3" width="130" height="41" align="center"><%=rsbanco("os_militarexec")%></td>
      <td class="fundo1" width="109" height="41" align="center"><b><font size="2">Data Conclusão</font></b></td>
      <td class="fundo3" width="141" height="41" align="center"><font size="2"><%=rsbanco("os_dataconc")%></td>
    </tr>
    <tr>
      <td class="fundo1" width="94" height="42" align="center"><b><font size="2">Ramal</font></b></td>
      <td class="fundo3" width="110" height="42" align="center"><%=rsbanco("os_ramalaber")%></td>
      <td class="fundo1" width="102" height="42" align="center"><b><font size="2">Tempo H/H</font></b></td>
      <td class="fundo3" width="130" height="42" align="center"><%=rsbanco("os_tempoexec")%></td>
      <td class="fundo1" width="109" height="42" align="center"><b><font size="2">Militar Entregou</font></b></td>
      <td class="fundo3" width="141" height="42" align="center"><%=rsbanco("os_militarconc")%></td>
    </tr>
    <tr>
      <td class="fundo1" width="94" height="42" align="center"><b><font size="2">Data Abertura</font></b></td>
      <td class="fundo3" width="110" height="42" align="center"><font size="2"><%=rsbanco("os_dataaber")%></td>
      <td class="fundo1" width="102" height="42" align="center"><b><font size="2">Data Execução</font></b></td>
      <td class="fundo3" width="130" height="42" align="center"><font size="2"><%=rsbanco("os_dataexec")%></td>
      <td class="fundo1" width="109" height="42" align="center"><b><font size="2">Militar Receb.</font></b></td>
      <td class="fundo3" width="141" height="42" align="center"><%=rsbanco("os_milrecconc")%></td>
    </tr>
    <tr>
      <td class="fundo1" width="206" height="94" colspan="2" rowspan="2" valign="middle" align="center"><b><font size="2">Material Utilizado</font></b></td>
      <td class="fundo3" width="234" height="91" colspan="2" rowspan="2" align="center"><%=rsbanco("os_matusadoexec")%></td>
      <td class="fundo1" width="109" height="36" align="center"><b><font size="1">Status da OS</font></b></td>
<% If rsbanco("os_status") = 1 then %>
	  <td class="fundo3" width="141" height="29" align="center"><img border="0" src="bolaverde.gif"></td>      
<% Else
		If rsbanco("os_status") = 2 then %>
      <td class="fundo3" width="141" height="29" align="center"><img border="0" src="bolaamarela.gif"></td>
		<% Else
			If rsbanco("os_status") = 3 then %>
      <td class="fundo3" width="141" height="29" align="center"><img border="0" src="bolaazul.gif"></td>		
			<% Else %>
      <td class="fundo3" width="141" height="29" align="center"><img border="0" src="bolavermelha.gif"></td>
			<% End If
		End If
	End If %>     
    </tr>
    <tr>
      <td class="fundo2" width="252" height="56" colspan="2">&nbsp;</td>
    </tr>
  </table>
<input type=hidden name=os_codigo value="<%=rsbanco("os_codigo")%>"><input type="submit" value="&nbsp;&nbsp;Excluir&nbsp;&nbsp;" name="BTincluir"></form>
<form action="consultasos2.asp"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;">
<table border="0" cellpadding="0" cellspacing="0" width="780" height="19">
    <tr>
      	<TD WIDTH=780 HEIGHT=19 COLSPAN=7 align="center" style="color:606060"><font color="#ffffff" size="1">Copyright (c) 2006. Hallyz Cia & Ltda. Todos os direitos reservados.</font></TD>
    </tr>
</table></form></center></body></html>
<% Else
	Response.Redirect("admin.asp")
End If %>