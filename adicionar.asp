<head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<body><center><form method="GET" action="add.asp">
<script language="JavaScript">
<!--
var months=new Array(13);
months[1]="Janeiro";
months[2]="Fevereiro";
months[3]="Mar&ccedil;o";
months[4]="Abril";
months[5]="Maio";
months[6]="Junho";
months[7]="Julho";
months[8]="Agosto";
months[9]="Setembro";
months[10]="Outubro";
months[11]="Novembro";
months[12]="Dezembro";
var time=new Date();
var lmonth=months[time.getMonth() + 1];
var date=time.getDate();
year=time.getFullYear();
var today = new Date();
var hrs = today.getHours();
document.write("<input name='os_dia' type='hidden' value='" + date + "' size='8' style='text-align: center'>");
document.write("<input name='os_mes' type='hidden' value='" + (time.getMonth() + 1) + "' size='8' style='text-align: center'>");
document.write("<input name='os_ano' type='hidden' value='" + year + "' size='8' style='text-align: center'>");
//-->
</script>
<table border="1" width="600" height="102">
    <tr>
      <td class="fundo1" width="90" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="504" height="102" align="center">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Controle de Ordem de Serviço</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Solicitação de 
      Abertura de Ordem de Serviço</b></font></td>
    </tr>
</table>
<table border="1" width="600" height="268">
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>Data</b></td>
        <td class="fundo3" width="200" height="23"><input name="os_solicdata" type="text" value="<% Response.write Now %>" size="20" style="border-style: inset; border-width: 5"></td>
        <td class="fundo1" width="100" height="23" align="center"><b>Periférico</b></td>
        <td class="fundo3" width="200" height="23" colspan="3">
	<% 	Set rsbanco5 = Server.CreateObject("ADODB.Recordset")
		rsbanco5.Open "SELECT * FROM periferico ORDER BY periferico_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="os_solicperiferico">
				<option value="0" selected>Selecione</option>								    
		<%  Do While Not rsbanco5.EOF %>								    
				<option value="<% = rsbanco5("periferico_nome") %>"><% = rsbanco5("periferico_nome") %></option>
		<% 		rsbanco5.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco5.Close
		Set rsbanco5= Nothing %>
  		</td>
      </tr>     
      <tr>
        <td class="fundo1" width="100" height="146" align="center"><b>Descrição</b></p><p align="center"><b>do</b></p><p align="center"><b>Problema</b></td>
        <td class="fundo3" width="500" height="146" colspan="3"><textarea name="os_solicdescricao" cols="60" rows="9" style="border-style: inset; border-width: 5">Se estiver solicitando Senha/Login favor informar:       Nome Completo, Nome de Guerra, Senha Sugerida e o        Sistema(E-mail, Rede, Internet...)</textarea></td>
      </tr>
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>Solicitante</b></td>
        <td class="fundo3" width="200" valign="middle" height="23"><input type="text" name="os_solicmilitar" size="20" style="border-style: inset; border-width: 5"></td>
        <td class="fundo1" width="100" height="23" align="center"><b>Esquadrão</b></td>
        <td class="fundo3" width="200" height="23">
	<% 	Set rsbanco4 = Server.CreateObject("ADODB.Recordset")
		rsbanco4.Open "SELECT * FROM esquadrao ORDER BY esquadrao_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="os_solicesquadrao">
				<option value="0" selected>Selecione</option>							    
		<%  Do While Not rsbanco4.EOF %>								    
				<option value="<% = rsbanco4("esquadrao_nome") %>"><% = rsbanco4("esquadrao_nome") %></option>
		<% 		rsbanco4.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco4.Close
		Set rsbanco4= Nothing %>
        </td>
      </tr>      
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>Seção</b></td>
        <td class="fundo3" width="500" height="23">
	<% 	Set rsbanco3 = Server.CreateObject("ADODB.Recordset")
		rsbanco3.Open "SELECT * FROM secao ORDER BY secao_nome ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="os_solicsecao">			
				<option value="0" selected>Selecione</option>					    
		<%  Do While Not rsbanco3.EOF %>								    
				<option value="<% = rsbanco3("secao_nome") %>"><% = rsbanco3("secao_nome") %></option>
		<% 		rsbanco3.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco3.Close
		Set rsbanco3= Nothing %> 
        </td>
        <td class="fundo1" width="100" height="23" align="center"><b>Ramal</b></td>
        <td class="fundo3" width="200" valign="middle" height="23"><input type="text" name="os_solicramal" size="12" style="border-style: inset; border-width: 5"></td>
      </tr>       
    </table>
<input type="submit" value="Salvar" name="BTincluir">&nbsp;&nbsp;&nbsp;&nbsp;<input type="reset" value="Limpar" name="BTlimpar"></form>
<form action="index.asp"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;">
<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2" color="#FFFFFF">OBS: Todos os Campos são obrigatórios. Verifique as informações antes de salvá-las,</font></b></p>
<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2" color="#FFFFFF">&nbsp; pois as mesmas não poderão ser apagadas pelo usuário solicitante.</font></b></p>
<!--#include file="rodape.asp"--></form>
</center></body></html>