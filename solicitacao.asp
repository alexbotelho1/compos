<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% codsolic = request.querystring("codsolic")

	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from os WHERE os_codigo = "&codsolic&"",banco,AdOpenKeySet,AdLockOptimistic %>
<body>
<center>
<table border="1" width="600" height="102">
    <tr>
      <td class="fundo1" width="100" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="500" height="102" align="center">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Controle de Ordem de Serviço</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Solicitação de abertura de Ordem de Serviço</b></font></td>
    </tr>
</table>
<table border="1" width="600" height="23">
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>Número</b></td>
        <td class="fundo3" width="200" height="23" align="center"><%=rsbanco("os_codigo")%></font></td>
        <td class="fundo1" width="100" height="23" align="center"><b>Data</b></td>
        <td class="fundo3" width="200" height="23" align="center"><%=rsbanco("os_solicdata")%></font></td>
      </tr>    
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>Solicitante</b></td>
        <td class="fundo3" width="200" height="23" align="center"><%=rsbanco("os_solicmilitar")%></font></td>
        <td class="fundo1" width="100" height="23" align="center"><b>Periférico</b></td>
        <td class="fundo3" width="200" height="23" align="center"><%=rsbanco("os_solicperiferico")%></font></td>
      </tr>
<% If rsbanco("os_solicperiferico") = "Login/Senha" Then
	If Session("Level") = 1 or Session("Level") = 2 Then %>
     <tr>
        <td class="fundo1" width="100" height="128" align="center"><b>Descrição</b></p><p align="center"><b>do</b></p><p align="center"><b>Problema</b></td>
        <td class="fundo3" width="500" height="128" colspan="3"><%=rsbanco("os_solicdescricao")%></font></td>
      </tr>
    <% Else %>
      <tr>
        <td class="fundo1" width="100" height="128" align="center"><b>Descrição</b></p><p align="center"><b>do</b></p><p align="center"><b>Problema</b></td>
        <td class="fundo3" width="500" height="128" colspan="3" align="center"><font color="#FF0000"><b>Campo bloqueado por motivo de conter informações restritas aos administradores.<br><br>Obrigado pela comprreensão e volte sempre!</b></font></td>
      </tr>
	<% End If 
Else %>
      <tr>
        <td class="fundo1" width="100" height="128" align="center"><b>Descrição</b></p><p align="center"><b>do</b></p><p align="center"><b>Problema</b></td>
        <td class="fundo3" width="500" height="128" colspan="3"><%=rsbanco("os_solicdescricao")%></font></td>
      </tr>
<% End If %>
   </table>
   <table border="1" width="600" height="23">   
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>Esquadrão</b></td>
        <td class="fundo3" width="100" height="23" align="center"><%=rsbanco("os_solicesquadrao")%></font></td>
        <td class="fundo1" width="100" height="23" align="center"><b>Seção</b></td>
        <td class="fundo3" width="150" height="23" align="center"><%=rsbanco("os_solicsecao")%></font></td>
        <td class="fundo1" width="100" height="23" align="center"><b>Ramal</b></td>
        <td class="fundo3" width="50" height="23" align="center"><%=rsbanco("os_solicramal")%></font></td>
      </tr>       
    </table>
<form action="javascript:history.go(-1)"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;">
<!--#include file="rodape.asp"--></form></center></body></html>