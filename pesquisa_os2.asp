<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<%  If Session("LoggedIn") = True Then
	
	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from os order by os_numero ASC",banco,AdOpenKeySet,AdLockOptimistic
	set rsbanco1=server.createobject("ADODB.Recordset")
		rsbanco1.open "Select * from os order by os_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic
		
	os_dia1=CInt(request.querystring("os_dia"))
	os_mes1=CInt(request.querystring("os_mes"))
	os_ano1=CInt(request.querystring("os_ano"))
	os_dia2=CInt(request.querystring("os_dia1"))
	os_mes2=CInt(request.querystring("os_mes1"))
	os_ano2=CInt(request.querystring("os_ano1"))

erro = 0
	If os_dia1 > os_dia2 Or os_mes1 > os_mes2 Or os_ano1 > os_ano2 then
		erro = erro + 1
	End If
	
	'rsbanco.PageSize = 30  %>
<body><center>
	<% If erro > 0 then %>
<center>
<table border="0" width="700" height="23" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
	<tr>
    	<td width="97" height="101" align="center" style="border-left: 1px solid #000000; border-right-width: 1; border-top: 1px solid #000000; border-bottom-width: 1"><img border="0" src="logo.gif" width="73" height="90"></td>
    	<td width="603" height="101" align="center" style="border-left: 1px solid #000000; border-right: 1px solid #000000; border-top: 1px solid #000000; border-bottom-width: 1"><p style="margin-top: 0; margin-bottom: 0"><img border="0" src="aeronautica2.jpg" width="60" height="50"></p>
        <p style="margin-top: 0; margin-bottom: 0"><b><font size="4">COMANDO DA AERON�UTICA</font></b></p>
        <p style="margin-top: 0; margin-bottom: 0">BASE A�REA DE PORTO VELHO</p>
        <p style="margin-top: 0; margin-bottom: 0"><i><font size="2">Se��o de Tecnologia da Informa��o</font></i></td>
    </tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="740">		  
	<tr>
		<td align="center"><font face="Trebuchet MS" size="2" color="#FFFFFF"><i>Voc� n�o informou algum intervalo de pesquisa! Volte e corrija esse problema (erro <% = erro %>).</i></font></td>
	</tr>
</table>
<form action="javascript:history.go(-1)"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;"></form></center>
	<% Else %>	
<table border="0" width="700" height="23" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
	<tr>
    	<td width="97" height="101" align="center" style="border-left: 1px solid #000000; border-right-width: 1; border-top: 1px solid #000000; border-bottom-width: 1"><img border="0" src="logo.gif" width="73" height="90"></td>
    	<td width="603" height="101" align="center" style="border-left: 1px solid #000000; border-right: 1px solid #000000; border-top: 1px solid #000000; border-bottom-width: 1"><p style="margin-top: 0; margin-bottom: 0"><img border="0" src="aeronautica2.jpg" width="60" height="50"></p>
        <p style="margin-top: 0; margin-bottom: 0"><b><font size="4">COMANDO DA AERON�UTICA</font></b></p>
        <p style="margin-top: 0; margin-bottom: 0">BASE A�REA DE PORTO VELHO</p>
        <p style="margin-top: 0; margin-bottom: 0"><i><font size="2">Se��o de Tecnologia da Informa��o</font></i></td>
    </tr>
<% If (rsbanco.BOF And rsbanco.EOF) Or rsbanco.PageCount = 0 Then %>
<table border="0" cellpadding="0" cellspacing="0" width="100%" align="center">		  
	<tr>
		<td align="center"><font face="Trebuchet MS" size="2" color="#ffffff"><i>N&atilde;o h&aacute; nada no momento!</i></font></td>
	</tr>
</table>
<form action="pesquisa_os.asp"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;"></form>
<% Else %>     
	<tr>
    	<td width="700" height="23" colspan="2" align="center" style="border-left: 1px solid #000000; border-right: 1px solid #000000; border-top: 1px solid #000000; border-bottom-width: 1"><i><b>Relat�rio Mensal das Ordem de Servi�os</b></i></td>
    </tr>    
    </table>
  			<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" bgcolor="#FFFFFF" width="700" height="23">
    			<tr>
      				<td class="fundo6" width="42" height="23" align="center">OS</td>
      				<td class="fundo6" width="52" height="23" align="center">Esq</td>
      				<td class="fundo6" width="108" height="23" align="center">Se��o</td>
      				<td class="fundo6" width="98" height="23" align="center">Perif�rico</td>
      				<td class="fundo6" width="123" height="23" align="center">Solicitador</td>
      				<td class="fundo6" width="142" height="23" align="center">Militar da STI</td>
      				<td class="fundo6" width="77" height="23" align="center">Status</td>
    			</tr>
<%	numeroos = 1

	rsbanco1.MoveLast
		ultimo = rsbanco1("os_codigo")
	rsbanco1.MoveFirst
		
	Do While rsbanco1("os_codigo") <> ultimo
		If rsbanco1("os_numero") > 0 then
			numeroos = numeroos + 1
			rsbanco1.MoveNext
		Else
			rsbanco1.MoveNext
		End If
	Loop 
	
HowMany = 0
Do While Not rsbanco.EOF And HowMany < 20
	cont1 = 0
	Do While Not rsbanco.EOF And cont1 < 1 
		If rsbanco("os_dia") > (os_dia1 - 1) And rsbanco("os_dia") < (os_dia2 + 1) And rsbanco("os_mes") > (os_mes1 - 1) And rsbanco("os_mes") < (os_mes2 + 1) And rsbanco("os_ano") > (os_ano1 - 1) And rsbanco("os_ano") < (os_ano2 + 1) then %>  			
    			<tr>
      				<td width="42" height="23" align="center"><font size="2"><% = rsbanco("os_numero") %></td>
      				<td width="52" height="23" align="center"><font size="2"><% = rsbanco("os_solicesquadrao") %></td>
      				<td width="108" height="23" align="center"><font size="2"><% = rsbanco("os_solicsecao") %></td>
      				<td width="98" height="23" align="center"><font size="2"><% = rsbanco("os_solicperiferico") %></td>
      				<td width="123" height="23" align="center"><font size="2"><% = rsbanco("os_solicmilitar") %></td>
      				<td width="142" height="23" align="center"><font size="2"><% = rsbanco("os_militaraber") %></td>
<% 			If rsbanco("os_status") = 0 then %>
      				<td width="77" height="23" align="center"><font size="2">N�o Aberta</font></td>
<% 			Else
				If rsbanco("os_status") = 1 then %>
      				<td width="77" height="23" align="center"><font size="2">Aberta</font></td>
<%				Else 	
					If rsbanco("os_status") = 2 then %>
      				<td width="77" height="23" align="center"><font size="2">Executada</font></td>
		<% 			Else %>
      				<td width="77" height="23" align="center"><font size="2">Fechada</font></td>
    	<% 			End If
    			End IF
    		End If %>
    			</tr>
<% 			HowMany = HowMany + 1
			cont1 = cont1 + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
	cont2 = 0
	Do While Not rsbanco.EOF And cont2 < 1 
		If rsbanco("os_dia") > (os_dia1 - 1) And rsbanco("os_dia") < (os_dia2 + 1) And rsbanco("os_mes") > (os_mes1 - 1) And rsbanco("os_mes") < (os_mes2 + 1) And rsbanco("os_ano") > (os_ano1 - 1) And rsbanco("os_ano") < (os_ano2 + 1) then %>  				
    			<tr>
      				<td class="fundo6" width="42" height="23" align="center"><font size="2"><% = rsbanco("os_numero") %></td>
      				<td class="fundo6" width="52" height="23" align="center"><font size="2"><% = rsbanco("os_solicesquadrao") %></td>
      				<td class="fundo6" width="108" height="23" align="center"><font size="2"><% = rsbanco("os_solicsecao") %></td>
      				<td class="fundo6" width="98" height="23" align="center"><font size="2"><% = rsbanco("os_solicperiferico") %></td>
      				<td class="fundo6" width="123" height="23" align="center"><font size="2"><% = rsbanco("os_solicmilitar") %></td>
      				<td class="fundo6" width="142" height="23" align="center"><font size="2"><% = rsbanco("os_militaraber") %></td>
<% 			If rsbanco("os_status") = 0 then %>
      				<td class="fundo6" width="77" height="23" align="center"><font size="2">N�o Aberta</font></td>
<% 			Else
				If rsbanco("os_status") = 1 then %>
      				<td class="fundo6" width="77" height="23" align="center"><font size="2">Aberta</font></td>
<%				Else 	
					If rsbanco("os_status") = 2 then %>
      				<td class="fundo6" width="77" height="23" align="center"><font size="2">Executada</font></td>
		<% 			Else %>
      				<td class="fundo6" width="77" height="23" align="center"><font size="2">Fechada</font></td>
    	<% 			End If
    			End IF
    		End If %>
    			</tr>
<% 			HowMany = HowMany + 1
			cont2 = cont2 + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
Loop %>			
  			</table>    	
    	</td>
    </tr>    
</table>
<% If HowMany = 0 Then %>
<center>
<table border="0" cellpadding="0" cellspacing="0" width="740">		  
	<tr>
		<td align="center"><font face="Trebuchet MS" size="2" color="#FFFFFF"><i>Pesquisa sem resultado! Volte e informe outros dados � pesquisa.</i></font></td>
	</tr>
</table>
<form action="javascript:history.go(-1)"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;"></form></center>
<% End If %>
<form action="javascript:history.go(-1)"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;"></form>
<% End If %><% End If %>
</center></body></html>
<% Else
	Response.Redirect("admin.asp")
End If %>