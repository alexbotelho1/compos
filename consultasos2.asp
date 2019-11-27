<meta http-equiv="refresh" content="180"><html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% If Session("LoggedIn") = True Then

	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from os order by os_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic
			contreg = 0
			Do While Not rsbanco.EOF
				If rsbanco("os_numero") <> 0 then
					contreg = contreg + 1					
				End IF
				rsbanco.movenext	
			Loop

	If Request.QueryString("Ordem") = "" or Request.QueryString("Ordem") = "1" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_numero ASC",banco,AdOpenKeySet,AdLockOptimistic
	End If
	If Request.QueryString("Ordem") = "2" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_numero DESC",banco,AdOpenKeySet,AdLockOptimistic
	End If
	If Request.QueryString("Ordem") = "3" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_dataaber ASC",banco,AdOpenKeySet,AdLockOptimistic
	End If
	If Request.QueryString("Ordem") = "4" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_dataaber DESC",banco,AdOpenKeySet,AdLockOptimistic
	End If	
	If Request.QueryString("Ordem") = "5" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_solicperiferico ASC",banco,AdOpenKeySet,AdLockOptimistic
	End If	
	If Request.QueryString("Ordem") = "6" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_solicperiferico DESC",banco,AdOpenKeySet,AdLockOptimistic
	End If	
	If Request.QueryString("Ordem") = "7" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_solicsecao ASC",banco,AdOpenKeySet,AdLockOptimistic
	End If
	If Request.QueryString("Ordem") = "8" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_solicsecao DESC",banco,AdOpenKeySet,AdLockOptimistic
	End If
	If Request.QueryString("Ordem") = "9" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_solicesquadrao ASC",banco,AdOpenKeySet,AdLockOptimistic
	End If	
	If Request.QueryString("Ordem") = "10" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_solicesquadrao DESC",banco,AdOpenKeySet,AdLockOptimistic
	End If	
	If Request.QueryString("Ordem") = "11" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_militaraber ASC",banco,AdOpenKeySet,AdLockOptimistic
	End If	
	If Request.QueryString("Ordem") = "12" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_militaraber DESC",banco,AdOpenKeySet,AdLockOptimistic
	End If
	If Request.QueryString("Ordem") = "13" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_ramalaber ASC",banco,AdOpenKeySet,AdLockOptimistic
	End If	
	If Request.QueryString("Ordem") = "14" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_ramalaber DESC",banco,AdOpenKeySet,AdLockOptimistic
	End If
	If Request.QueryString("Ordem") = "15" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_status ASC",banco,AdOpenKeySet,AdLockOptimistic
	End If	
	If Request.QueryString("Ordem") = "16" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from os order by os_status DESC",banco,AdOpenKeySet,AdLockOptimistic
	End If
		
	rsbanco.PageSize = registros
	contpage = (contreg \ registros) + 1 %>
<body><center>
<table border="1" width="750" height="102">
    <tr>
      <td class="fundo1" width="90" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td width="660" height="102" align="center" class="fundo2">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Controle de Ordem de Serviço</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Consulta das Ordens de Serviço</b></font>
	  <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
	  <p style="margin-top: 0; margin-bottom: 0"><font color="#000000"><b>
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
document.write("<leftt>");
document.write("<B>");
document.write("<font face=Verdana size=1 color=#000000>");
if (hrs < 6)
document.write("Bom Dia!!! -");
else if (hrs < 12)
document.write("Bom Dia!!! -");
else if (hrs < 18)
document.write("Boa Tarde!!! -");
else
document.write("Boa Noite!!! -");   
document.write(" Porto Velho, ");
document.write(date + "  de ");
document.write(lmonth + " de  " + year + "</left></B>");
//-->
	  </script>	  
      </b></font></p></td>      
    </tr>
</table>
<% If (rsbanco.BOF And rsbanco.EOF) Or rsbanco.PageCount = 0 Then %>
<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">		  
	<tr>
		<td align="center"><font face="Trebuchet MS" size="2" color="#ffffff"><i>N&atilde;o h&aacute; nada no momento!</i></font></td>
	</tr>
</table>
<form action="cad_os.asp"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;"></form>
<% Else
		If Request.QueryString("Ordem") = "" Then
			Disposicao = 1
		Else 
			Disposicao = Request.QueryString("Ordem")
		End If
		If Request.QueryString("page") > 0 Then  			   
			rsbanco.AbsolutePage = Request.QueryString("page")					
		Else				
			rsbanco.AbsolutePage = 1					
		End If 
if Request.QueryString("page") = 1 Then %>	
<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">		  
	<tr>
		<td align="center" width="200">
		<% If rsbanco.AbsolutePage > 1 Then %>
			<a href="consultasos2.asp?Ordem=<% = Disposicao %>&page=<% = (rsbanco.AbsolutePage - 1) %>"><img src="paganterior2.gif" align="absmiddle" border="0"></a>
		<% Else %>
			<img src="paganterior1.gif" align="absmiddle">
		<% End If %>
		</td>
		<td align="center" width="350"><font face="Trebuchet MS" size="2" color="#FFFFFF"><b>Esses são os <% = (rsbanco.PageSize * rsbanco.AbsolutePage) %> primeiros</b></font></td>
		<td align="center" width="200">
		<% If rsbanco.AbsolutePage < contpage Then %>
      		<a href="consultasos2.asp?Ordem=<% = Disposicao %>&page=<% = (rsbanco.AbsolutePage + 1) %>"><img src="pagseguinte2.gif" align="absmiddle" border="0"></a>
   		<% 	Else %>
          	<img src="pagseguinte1.gif" align="absmiddle">
    	<% 	End If %>   			    		
       	</td>
	</tr>
</table>
<% Else %>
<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">		  
	<tr>
		<td align="center" width="200">
		<% If rsbanco.AbsolutePage > 1 Then %>
			<a href="consultasos2.asp?Ordem=<% = Disposicao %>&page=<% = (rsbanco.AbsolutePage - 1) %>"><img src="paganterior2.gif" align="absmiddle" border="0"></a>
		<% Else %>
			<img src="paganterior1.gif" align="absmiddle">
		<% End If %>
		</td>
		<td align="center" width="350"><font face="Trebuchet MS" size="2" color="#FFFFFF"><b>Registros do <% = ((rsbanco.PageSize * rsbanco.AbsolutePage) - 10) %>° até <% = (rsbanco.PageSize * rsbanco.AbsolutePage) %>°</b></font></td>
		<td align="center" width="200">
		<% If rsbanco.AbsolutePage < contpage Then %>
      		<a href="consultasos2.asp?Ordem=<% = Disposicao %>&page=<% = (rsbanco.AbsolutePage + 1) %>"><img src="pagseguinte2.gif" align="absmiddle" border="0"></a>
   		<% 	Else %>
          	<img src="pagseguinte1.gif" align="absmiddle">
    	<% 	End If %>   			    		
       	</td>
	</tr>
</table>
<% End If
	pagina = rsbanco.AbsolutePage
	contador = contpage
%>
<!--#include file="statusos.asp"-->
<table width="750" align="center" border="1">
<tr class="fundo1">
	  <td width="80" height="14" align="center"><a href="consultasos2.asp?Ordem=1&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta.gif" alt="Crescente"></a>&nbsp;&nbsp;<a href="consultasos2.asp?Ordem=2&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta2.gif" alt="Decrescente"></a></td>
      <td width="140" height="14" align="center"><a href="consultasos2.asp?Ordem=3&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta.gif" alt="Crescente"></a>&nbsp;&nbsp;<a href="consultasos2.asp?Ordem=4&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta2.gif" alt="Decrescente"></a></td>
	  <td width="110" height="14" align="center"><a href="consultasos2.asp?Ordem=5&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta.gif" alt="Crescente"></a>&nbsp;&nbsp;<a href="consultasos2.asp?Ordem=6&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta2.gif" alt="Decrescente"></a></td>
	  <td width="130" height="14" align="center"><a href="consultasos2.asp?Ordem=7&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta.gif" alt="Crescente"></a>&nbsp;&nbsp;<a href="consultasos2.asp?Ordem=8&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta2.gif" alt="Decrescente"></a></td>
      <td width="70" height="14" align="center"><a href="consultasos2.asp?Ordem=9&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta.gif" alt="Crescente"></a>&nbsp;&nbsp;<a href="consultasos2.asp?Ordem=10&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta2.gif" alt="Decrescente"></td>
      <td width="130" height="14" align="center"><a href="consultasos2.asp?Ordem=11&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta.gif" alt="Crescente"></a>&nbsp;&nbsp;<a href="consultasos2.asp?Ordem=12&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta2.gif" alt="Decrescente"></td>
      <td width="50" height="14" align="center"><a href="consultasos2.asp?Ordem=13&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta.gif" alt="Crescente"></a>&nbsp;&nbsp;<a href="consultasos2.asp?Ordem=14&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta2.gif" alt="Decrescente"></td>
	  <td width="40" height="14" align="center"><a href="consultasos2.asp?Ordem=15&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta.gif" alt="Crescente"></a>&nbsp;&nbsp;<a href="consultasos2.asp?Ordem=16&page=<% = (rsbanco.AbsolutePage) %>"><img border="0" src="seta2.gif" alt="Decrescente"></td>
	  <td width="70" height="14" align="center">&nbsp;</td> 	    
</tr>
	<tr class="fundo1">
	  <td width="80" height="29" align="center"><font size="2"><b>N° OS</b></font></td>
      <td width="140" height="29" align="center"><font size="2"><b>Data Abertura</b></font></td>
      <td width="110" height="29" align="center"><b><font size="2">Periférico</td>
      <td width="130" height="29" align="center"><font size="2"><b>Seção</b></font></td>      
      <td width="70" height="29" align="center"><b><font size="2">Esquadrão</td>  
      <td width="130" height="29" align="center"><b><font size="2">Militar STI</td>
      <td width="50" height="29" align="center"><b><font size="2">Ramal</td>
	  <td width="40" height="29" align="center"><b><font size="2">Status</td>
	  <td width="70" height="29" align="center"><b><font size="2">Opções</td> 	      
</tr>
<% 	pagina1 = Request.QueryString("page")
 	If pagina1 > 1 then
		pulacont = 0
		Do While Not rsbanco.EOF And pulacont < registros
			If rsbanco("os_numero") <> 0 then
				pulacont = pulacont + 1
			End IF
			rsbanco.movenext
		Loop
	End If
	HowMany = 0
	Do While Not rsbanco.EOF And HowMany < rsbanco.PageSize
	If rsbanco("os_numero") <> 0 then %>
<tr class="fundo5">
	  <td width="80" height="29" align="center"><a href="os.asp?codsolic=<% = rsbanco("os_codigo") %>"><font color="#7F0D11" size="2"><b><% = rsbanco("os_numero") %></a></td>
      <td width="140" height="29" align="center"><font size="2"><% = rsbanco("os_dataaber") %></td>
      <td width="110" height="29"  align="center"><font size="2"><% = rsbanco("os_solicperiferico") %></td>
      <td width="130" height="29"  align="center"><font size="2"><% = rsbanco("os_solicsecao") %></td>       
      <td width="70" height="29"  align="center"><font size="2"><% = rsbanco("os_solicesquadrao") %></td> 
      <td width="130" height="29" align="center"><font size="2"><% = rsbanco("os_militaraber") %></td>
      <td width="50" height="29" align="center"><font size="2"><% = rsbanco("os_ramalaber") %></td>
<% If rsbanco("os_status") = 1 then %>
	  <td width="40" height="29" align="center"><img border="0" src="bolaverde.gif"></td>      
<% Else
		If rsbanco("os_status") = 2 then %>
      <td width="40" height="29" align="center"><img border="0" src="bolaamarela.gif"></td>
		<% Else
			If rsbanco("os_status") = 3 then %>
      <td width="40" height="29" align="center"><img border="0" src="bolaazul.gif"></td>		
			<% Else %>
      <td width="40" height="29" align="center"><img border="0" src="bolavermelha.gif"></td>
			<% End If
		End If
	End If %>
	  <td width="70" height="29" align="right"><font size="2">
	  	<% If rsbanco("os_status") = 1 then %><a href="os_executar.asp?codsolic=<% = rsbanco("os_codigo") %>"><img border="0" src="executar.gif" alt="Executar a Ordem de Serviço"></a>&nbsp;&nbsp;<% End If %>
	  	<% If rsbanco("os_status") = 2 then %><a href="os_fechar.asp?codsolic=<% = rsbanco("os_codigo") %>"><img border="0" src="fechar.gif" alt="Fechar a Ordem de Serviço"></a>&nbsp;&nbsp;<% End If %>
		<% If rsbanco("os_status") = 3 then %><img border="0" src="finalizada.gif" alt="Ordem de Serviço Finalizada"></a>&nbsp;&nbsp;<% End If %>	  
	  	<% If rsbanco("os_status") = 1 then %><a href="alt_os1.asp?codsolic=<% = rsbanco("os_codigo") %>"><img border="0" src="editar.gif" alt="Editar a Ordem de Serviço"></a>&nbsp;&nbsp;<% End If %>
	  	<% If rsbanco("os_status") = 2 then %><a href="alt_os2.asp?codsolic=<% = rsbanco("os_codigo") %>"><img border="0" src="editar.gif" alt="Editar a Ordem de Serviço"></a>&nbsp;&nbsp;<% End If %>
	  	<% If rsbanco("os_status") = 3 then %><a href="alt_os.asp?codsolic=<% = rsbanco("os_codigo") %>"><img border="0" src="editar.gif" alt="Editar a Ordem de Serviço"></a>&nbsp;&nbsp;<% End If %>
	  	<% If Session("Level") < 3 Then %><a href="exc_os.asp?codsolic=<% = rsbanco("os_codigo") %>"><img border="0" src="del.gif" alt="Excluir a Ordem de Serviço"></a><% End If %></td>
	</tr>
<%		HowMany = HowMany + 1			
	End IF		
	rsbanco.movenext			 
	Loop %>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">		  
	<tr>
		<td align="center" width="200">
		<% If pagina > 1 Then %>
			<a href="consultasos2.asp?Ordem=<% = Disposicao %>&page=<% = (pagina - 1) %>"><img src="paganterior2.gif" align="absmiddle" border="0"></a>
		<% Else %>
			<img src="paganterior1.gif" align="absmiddle">
		<% End If %>
		</td>
		<form action="cad_os.asp"><td align="center" width="350"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;"></td></form>
		<td align="center" width="200">
		<% If pagina < contador Then %>
      		<a href="consultasos2.asp?Ordem=<% = Disposicao %>&page=<% = (pagina + 1) %>"><img src="pagseguinte2.gif" align="absmiddle" border="0"></a>
   		<% 	Else %>
          	<img src="pagseguinte1.gif" align="absmiddle">
    	<% 	End If %>    			    		
       	</td>
	</tr>
</table>
<!--#include file="rodape.asp"-->
<% End If %>
</center></body></html>
<% Else
	Response.Redirect("admin.asp")
End If %>