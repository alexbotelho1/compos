<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% codsolic = request.querystring("codsolic")

	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from os WHERE os_codigo = "&codsolic&"",banco,AdOpenKeySet,AdLockOptimistic
%>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
<body><center>
<table border="1" width="700" height="102">
    <tr>
      <td class="fundo1" width="90" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td width="610" height="102" align="center" class="fundo2">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Controle de Ordem de Servi�o</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Consulta das solicita��es de abertura de Ordem de Servi�o</b></font></td>
    </tr>
</table>
<table border="1" width="700" height="410">
    <tr>
      <td class="fundo1" width="94" height="34" align="center"><b><font size="2">Os N�mero</font></b></td>
      <td class="fundo3" width="110" height="34" align="center"><%=rsbanco("os_numero")%></td>
      <td class="fundo1" width="102" height="34" align="center"><b><font size="2">Data Solicita��o</font></b></td>
      <td class="fundo3" width="130" height="34" align="center"><font size="2"><%=rsbanco("os_solicdata")%></td>
      <td class="fundo1" width="109" height="34" align="center"><b><font size="2">Solicitado Por</font></b></td>
      <td class="fundo3" width="141" height="34" align="center"><%=rsbanco("os_solicmilitar")%></td>
    </tr>
    <tr>
      <td class="fundo1" width="206" height="26" colspan="2" align="center"><b><font size="2">Observa��es (Abertura)</font></b></td>
      <td class="fundo1" width="234" height="26" colspan="2" align="center"><b><font size="2">Descri��o do Servi�o</font></b></td>
      <td class="fundo1" width="252" height="26" colspan="2" align="center"><b><font size="2">Observa��es (Conclus�o)</font></b></td>
    </tr>
    <tr>
      <td class="fundo3" width="206" height="121" colspan="2"><%=rsbanco("os_descricaoaber")%></td>
      <td class="fundo3" width="234" height="121" colspan="2"><%=rsbanco("os_descricaoexec")%></td>
      <td class="fundo3" width="252" height="121" colspan="2"><%=rsbanco("os_observconc")%></td>
    </tr>
    <tr>
      <td class="fundo1" width="94" height="41" align="center"><b><font size="2">Militar STI</font></b></td>
      <td class="fundo3" width="110" height="41" align="center"><%=rsbanco("os_militaraber")%></td>
      <td class="fundo1" width="102" height="41" align="center"><b><font size="2">Militar Exec.</font></b></td>
      <td class="fundo3" width="130" height="41" align="center"><%=rsbanco("os_militarexec")%></td>
      <td class="fundo1" width="109" height="41" align="center"><b><font size="2">Data Conclus�o</font></b></td>
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
      <td class="fundo1" width="102" height="42" align="center"><b><font size="2">Data Execu��o</font></b></td>
      <td class="fundo3" width="130" height="42" align="center"><font size="2"><%=rsbanco("os_dataexec")%></td>
      <td class="fundo1" width="109" height="42" align="center"><b><font size="2">Militar Receb.</font></b></td>
      <td class="fundo3" width="141" height="42" align="center"><%=rsbanco("os_milrecconc")%></td>
    </tr>
    <tr>
      <td class="fundo1" width="206" height="94" colspan="2" rowspan="2" valign="middle" align="center"><b><font size="2">Material Utilizado</font></b></td>
      <td class="fundo3" width="234" height="91" colspan="2" rowspan="2" align="center"><%=rsbanco("os_matusadoexec")%></td>
      <td class="fundo1" width="109" height="36" align="center"><b><font size="1">Status da OS</font></b></td>
<% If rsbanco("os_status") = 1 then %>
	  <td class="fundo1" width="141" height="29" align="center"><img border="0" src="bolaverde.gif"></td>      
<% Else
		If rsbanco("os_status") = 2 then %>
      <td class="fundo1" width="141" height="29" align="center"><img border="0" src="bolaamarela.gif"></td>
		<% Else
			If rsbanco("os_status") = 3 then %>
      <td class="fundo1" width="141" height="29" align="center"><img border="0" src="bolaazul.gif"></td>		
			<% Else %>
      <td class="fundo1" width="141" height="29" align="center"><img border="0" src="bolavermelha.gif"></td>
			<% End If
		End If
	End If %>     
    </tr>
    <tr>
      <td class="fundo5" width="252" height="56" align="center" colspan="2"><input type=button onclick="MM_openBrWindow('imprimir_os.asp?codsolic=<%=rsbanco("os_codigo")%>','','width=620,height=500,scrollbars=yes,menubar=yes')" value=" Imprimir " style="background-position: center 50%; font-family:ADMUI3Sm; font-size:8 pt; color:#000000; font-weight:bold; background-repeat:no-repeat"></td>
    </tr>
</table>
<form action="javascript:history.go(-1)"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;">
<!--#include file="rodape.asp"--></form></center></body></html>