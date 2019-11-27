<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% 	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from os",banco,AdOpenKeySet,AdLockOptimistic
	set rsbanco1=server.createobject("ADODB.Recordset")
		rsbanco1.open "Select * from os order by os_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic
		
	os_dia=CInt(request.querystring("os_dia"))
	os_mes=CInt(request.querystring("os_mes"))
	os_ano=CInt(request.querystring("os_ano"))
	os_dia1=CInt(request.querystring("os_dia1"))
	os_mes1=CInt(request.querystring("os_mes1"))
	os_ano1=CInt(request.querystring("os_ano1"))
	
	'rsbanco.PageSize = 30  %>
<body topmargin="0" leftmargin="0"><center>
	<% If os_dia = 0 And os_mes = 0 And os_ano = 0 And os_dia1 = 0 And os_mes1 = 0 And os_ano1 = 0 then %>
<center>
<table border="0" cellpadding="0" cellspacing="0" width="740">		  
	<tr>
		<td align="center"><font face="Trebuchet MS" size="2" color="#FFFFFF"><i>Você não informou algum intervalo de pesquisa! Volte e corrija esse problema.</i></font></td>
	</tr>
</table>
<form action="javascript:history.go(-1)"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;"></form></center>
	<% Else %>	
<table border="0" width="700" height="23" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
	<tr>
    	<td width="97" height="101" align="center" style="border-left: 1px solid #000000; border-right-width: 1; border-top: 1px solid #000000; border-bottom-width: 1"><img border="0" src="logo.gif" width="73" height="90"></td>
    	<td width="603" height="101" align="center" style="border-left: 1px solid #000000; border-right: 1px solid #000000; border-top: 1px solid #000000; border-bottom-width: 1"><p style="margin-top: 0; margin-bottom: 0"><img border="0" src="aeronautica2.jpg" width="60" height="50"></p>
        <p style="margin-top: 0; margin-bottom: 0"><b><font size="4">COMANDO DA AERONÁUTICA</font></b></p>
        <p style="margin-top: 0; margin-bottom: 0">BASE AÉREA DE PORTO VELHO</p>
        <p style="margin-top: 0; margin-bottom: 0"><i><font size="2">Seção de Tecnologia da Informação</font></i></td>
    </tr>
	<tr>
    	<td width="700" height="23" colspan="2" align="center" style="border-left: 1px solid #000000; border-right: 1px solid #000000; border-top: 1px solid #000000; border-bottom-width: 1"><i><b>Relatório Mensal das Ordem de Serviços</b></i></td>
    </tr>    
    </table>
  			<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" bgcolor="#FFFFFF" width="700" height="23">
    			<tr>
      				<td class="fundo6" width="42" height="23" align="center">OS</td>
      				<td class="fundo6" width="52" height="23" align="center">Esq</td>
      				<td class="fundo6" width="108" height="23" align="center">Seção</td>
      				<td class="fundo6" width="98" height="23" align="center">Periférico</td>
      				<td class="fundo6" width="123" height="23" align="center">Solicitador</td>
      				<td class="fundo6" width="142" height="23" align="center">Militar 
                    da STI</td>
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
		If rsbanco("os_mes") = os_mes then %>  			
    			<tr>
      				<td width="42" height="23" align="center"><font size="2"><% = rsbanco("os_numero") %></td>
      				<td width="52" height="23" align="center"><font size="2"><% = rsbanco("os_solicesquadrao") %></td>
      				<td width="108" height="23" align="center"><font size="2"><% = rsbanco("os_solicsecao") %></td>
      				<td width="98" height="23" align="center"><font size="2"><% = rsbanco("os_solicperiferico") %></td>
      				<td width="123" height="23" align="center"><font size="2"><% = rsbanco("os_solicmilitar") %></td>
      				<td width="142" height="23" align="center"><font size="2"><% = rsbanco("os_militaraber") %></td>
<% 			If rsbanco("os_status") = 0 then %>
      				<td width="77" height="23" align="center"><font size="2">Não Aberta</font></td>
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
		If rsbanco("os_mes") = os_mes then %>     				
    			<tr>
      				<td class="fundo6" width="42" height="23" align="center"><font size="2"><% = rsbanco("os_numero") %></td>
      				<td class="fundo6" width="52" height="23" align="center"><font size="2"><% = rsbanco("os_solicesquadrao") %></td>
      				<td class="fundo6" width="108" height="23" align="center"><font size="2"><% = rsbanco("os_solicsecao") %></td>
      				<td class="fundo6" width="98" height="23" align="center"><font size="2"><% = rsbanco("os_solicperiferico") %></td>
      				<td class="fundo6" width="123" height="23" align="center"><font size="2"><% = rsbanco("os_solicmilitar") %></td>
      				<td class="fundo6" width="142" height="23" align="center"><font size="2"><% = rsbanco("os_militaraber") %></td>
<% 			If rsbanco("os_status") = 0 then %>
      				<td class="fundo6" width="77" height="23" align="center"><font size="2">Não Aberta</font></td>
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
		<td align="center"><font face="Trebuchet MS" size="2" color="#FFFFFF"><i>Você não informou algum intervalo de pesquisa! Volte e corrija esse problema. <% = os_mes %></i></font></td>
	</tr>
</table>
<form action="javascript:history.go(-1)"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;"></form></center>
<% End If %>		
<form action="javascript:history.go(-1)"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;"></form>
<% End If %>
</center></body></html>